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

        public void ApplyCheaterExpertiese(string resultPath)
        {
            Console.WriteLine("Begin data loading");

            context = new OidzDbContext();
            context.ChangeTracker.AutoDetectChangesEnabled = false;
            context.Database.SetCommandTimeout((int)TimeSpan.FromMinutes(30).TotalSeconds);

            int amount = 10000;
            int count = context.Users.Count();
            double cyclesCount = count / amount;
            List<CheaterModel> cheaters = new List<CheaterModel>();

            for (double i = 0; i < cyclesCount; i++)
            {
                var users = context.Users
                .Skip((int)i * amount)
                .Take(amount)
                .Include(u => u.Events.Where(e => e.Type == 6 || e.Type == 5 || e.Type == 4))
                .ThenInclude(e => e.CurrencyPurchase)
                .Include(u => u.Events.Where(e => e.Type == 6 || e.Type == 5 || e.Type == 4))
                .ThenInclude(e => e.ItemPurchase)
                .Include(u => u.Events.Where(e => e.Type == 6 || e.Type == 5 || e.Type == 4))
                .ThenInclude(e => e.StageEnd)
                .ToList()
                ;

                cheaters.AddRange(users.Select(GetCheaterModel)
                    .Where(m => m.CurrecyRecieved < m.CurrencySpent));

                Console.WriteLine($"{i} / {cyclesCount} cycles ended");

                context = new OidzDbContext();
                context.ChangeTracker.AutoDetectChangesEnabled = false;

            }

            JsonSerializer serializer = new JsonSerializer();

            using (StreamWriter file = File.CreateText(resultPath + "\\CheatersCache.json"))
            {
                serializer.Serialize(file, cheaters);
            }
        }

        public void UploadClusteringResults(string path)
        {
            JsonSerializer serializer = new JsonSerializer();
            List<UserClusteringModel> rawData;

            Console.WriteLine("Reading data from file...");
            using (StreamReader file = File.OpenText(path + "//CheatersCache.json"))
            {
                rawData = (List<UserClusteringModel>)serializer.Deserialize(file, typeof(List<UserClusteringModel>));
            }

            Console.WriteLine("Data read!");
            Console.WriteLine(DateTime.Now);
            Console.WriteLine("Updating db...");

            UpdateClusters(rawData, 500);
        }

        public void UploadCheatersResults(string path)
        {
            JsonSerializer serializer = new JsonSerializer();
            List<CheaterModel> rawData;

            Console.WriteLine("Reading data from file...");
            using (StreamReader file = File.OpenText(path + "\\CheatersCache.json"))
            {
                rawData = (List<CheaterModel>)serializer.Deserialize(file, typeof(List<CheaterModel>));
            }

            Console.WriteLine("Data read!");
            Console.WriteLine(DateTime.Now);
            Console.WriteLine("Updating db...");

            UpdateCheaters(rawData, 500);
        }

        private void UpdateClusters(List<UserClusteringModel> rawData, int commitCount)
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
                        context = UpdateUserCluster(context, model, count, commitCount, true);
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

        private void UpdateCheaters(List<CheaterModel> rawData, int commitCount)
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
                        context = UpdateUserIsCheater(context, model, count, commitCount, true);
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

        private OidzDbContext UpdateUserCluster(
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

        private OidzDbContext UpdateUserIsCheater(
            OidzDbContext context,
            CheaterModel model,
            int count,
            int commitCount,
            bool recreateContext)
        {

            var user = context.Users.Find(model.Id);
            user.IsCheater = true;

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

        private CheaterModel GetCheaterModel(User user)
        {
            int recieved = 0;
            int spent = 0;

            foreach (var ev in user.Events)
            {
                switch(ev.Type) {
                    case 6:
                        recieved += (int)ev.CurrencyPurchase.Currency;
                        break;

                    case 5:
                        spent += (int)ev.ItemPurchase.Price;
                        break;

                    case 4:
                        recieved += (int)ev.StageEnd.Currency;
                        break;
                }
            }

            return new CheaterModel() 
            { 
                Id = user.Id,
                CurrecyRecieved = recieved,
                CurrencySpent = spent,
            };
        }
    }
}
