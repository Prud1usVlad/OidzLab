using DataAcquisition.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataAcquisition.Models; 


namespace DataAcquisition.Util
{
    internal class RepetitionsRemover
    {
        private List<EventViewModel> rawData = new List<EventViewModel>();
        private SortedSet<string> userGuids = new SortedSet<string>();

        public void RemoveRepetitions(string path, string resultPath, int counter)
        {
            Console.WriteLine("File " + counter + " started");

            JsonSerializer serializer = new JsonSerializer();

            using (StreamReader file = File.OpenText(path))
            {
                rawData = (List<EventViewModel>)serializer.Deserialize(file, typeof(List<EventViewModel>));
            }

            var result = new List<EventViewModel>();

            foreach (var item in rawData)
            {
                if (item.Event_id == 2)
                {
                    if (!userGuids.Add(item.Udid.ToString()))
                        continue;
                }

                result.Add(item);
            }

            //open file stream
            using (StreamWriter file = File.CreateText(resultPath + "\\file_" + counter + ".json"))
            {
                serializer.Serialize(file, result);
            }

            Console.WriteLine("File " + counter + " processed");
            Console.WriteLine("Current amount of users: " + userGuids.Count());
        }
    }
}
