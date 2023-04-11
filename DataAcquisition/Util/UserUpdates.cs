using DataAcquisition.Models;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Npgsql.Replication.PgOutput.Messages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAcquisition.Util
{
    public class UserUpdates
    {
        private OidzDbContext context;

        // Clustering results are in large Results/ClusteringCache.json
        public void ApplyClustering(string resultPath)
        {
            Console.WriteLine("Begin data loading");

            context = new OidzDbContext();
            context.ChangeTracker.AutoDetectChangesEnabled = false;
            context.Database.SetCommandTimeout((int)TimeSpan.FromMinutes(30).TotalSeconds);


            var users = context.Users
                .Include(u => u.Events.Where(e => e.Type == 6))
                .ThenInclude(e => e.CurrencyPurchase)
                .ToList();

            double[][] rawData = users
                .Select(u => new double[] { (double)u.Events.Sum(e => e.CurrencyPurchase.Price) })
                .Where(a => a[0] != 0)
                .ToArray();

            double[,] data = new double[rawData.Length, 1];

            for (int i = 0; i < rawData.Length; i++)
            {
                data[i, 0] = rawData[i][0];
            }


            Console.WriteLine("-------------------");
            Console.WriteLine("Begin k-means clustering");

            int numClusters = 3;

            alglib.clusterizerstate s;
            alglib.kmeansreport rep;

            alglib.clusterizercreate(out s);
            alglib.clusterizersetpoints(s, data, 2);
            alglib.clusterizersetkmeanslimits(s, 5, 0);
            alglib.clusterizerrunkmeans(s, 3, out rep);


            var toJson = new List<UserClusteringModel>();

            for (int i = 0; i < data.Length; i++)
            {
                toJson.Add(new UserClusteringModel { Cluster = rep.cidx[i] + 1, Id = users[i].Id, Value = rawData[i][0] });
            }


            JsonSerializer serializer = new JsonSerializer();

            using (StreamWriter file = File.CreateText(resultPath + "\\ClusteringCache.json"))
            {
                serializer.Serialize(file, toJson);
            }

        }


        public void UploadClusteringResults(string path)
        {
            JsonSerializer serializer = new JsonSerializer();
            List<UserClusteringModel> rawData;

            Console.WriteLine("Reading data from file...");
            using (StreamReader file = File.OpenText(path))
            {
                rawData = (List<UserClusteringModel>)serializer.Deserialize(file, typeof(List<UserClusteringModel>));
            }

            Console.WriteLine("Data read!");
            Console.WriteLine(DateTime.Now);
            Console.WriteLine("Updating db...");

            UpdateUserClusters(rawData, 500);



        }

        private void UpdateUserClusters(List<UserClusteringModel> rawData, int commitCount)
        {
            using (var ts = EtlCore.CreateTransactionScope(TimeSpan.FromMinutes(60)))
            {
                context = null;
                try
                {
                    context = new OidzDbContext();
                    context.ChangeTracker.AutoDetectChangesEnabled = false;

                    int count = 0;
                    foreach (var model in rawData)
                    {
                        context = UpdateUser(context, model, count, commitCount, true);
                        ++count;
                    }

                    context.SaveChanges();
                }
                finally
                {
                    if (context != null)
                        context.Dispose();
                }


                ts.Complete();
            }
        }

        private OidzDbContext UpdateUser(
            OidzDbContext context,
            UserClusteringModel model,
            int count,
            int commitCount,
            bool recreateContext)
        {

            var user = context.Users.Find(model.Id);
            user.Cluster = model.Cluster;

            context.Update(user);

            if (count % commitCount == 0)
            {
                context.SaveChanges();
                if (count != 0)
                {
                    Console.SetCursorPosition(0, Console.CursorTop - 1);
                    Program.ClearCurrentConsoleLine();
                }
                Console.WriteLine(count + " pieces processed. ");


                if (recreateContext)
                {
                    context.Dispose();
                    context = new OidzDbContext();
                    context.ChangeTracker.AutoDetectChangesEnabled = false;
                }
            }

            return context;
        }
    }
}
