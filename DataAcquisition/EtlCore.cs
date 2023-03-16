using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataAcquisition.Models.DataModels;
using DataAcquisition.Models;
using Microsoft.EntityFrameworkCore.Diagnostics;
using Newtonsoft.Json;



namespace DataAcquisition
{
    public class EtlCore
    {
        private List<EventViewModel> rawData;
        private PostgresContext context;

        public EtlCore() 
        { 
            context = new PostgresContext();
        }

        public void ReadData(string path)
        {
            using (StreamReader file = File.OpenText(path))
            {
                JsonSerializer serializer = new JsonSerializer();
                rawData = (List<EventViewModel>)serializer.Deserialize(file, typeof(List<EventViewModel>));
            }

            // rawData.ForEach(i => SaveData(i));
            var counter = 0;

            foreach (var piece in rawData)
            {
                SaveData(piece);

                if (++counter % 100 == 0)
                {
                    context.SaveChanges();
                    Console.Clear();
                    Console.WriteLine(counter + " pieces processed. " + (rawData.Count() - counter) + " to go..." );
                }
            }
                
        }

        private void SaveData(EventViewModel data)
        {
            switch (data.Event_id)
            {
                case 1:
                    SaveLaunch(data); 
                    break;
                case 2:
                    SaveFirstLaunch(data); 
                    break;
                case 3:
                    SaveStageStart(data); 
                    break;
                case 4:
                    SaveStageEnd(data); 
                    break;
                case 5:
                    SaveItemPurchase(data); 
                    break;
                case 6:
                    SaveCurrencyPurchase(data); 
                    break;
            }
        }

        private void SaveLaunch(EventViewModel eventVm)
        {
            var e = GetNewEvent(eventVm);

            context.Events.Add(e);
        }

        private void SaveFirstLaunch(EventViewModel eventVm)
        {
            var e = GetNewEvent(eventVm);
            var user = new User
            {
                Id = eventVm.Udid,
                Gender = eventVm.Parameters["gender"],
                Age = int.Parse(eventVm.Parameters["age"]),
                Country = eventVm.Parameters["country"],
            };

            context.Events.Add(e);
            context.Users.Add(user);
        }

        private void SaveCurrencyPurchase(EventViewModel eventVm)
        {
            var e = GetNewEvent(eventVm);
            var purchase = new CurrencyPurchase
            {
                Id = e.Id,
                PackName = eventVm.Parameters["name"],
                Price = decimal.Parse(eventVm.Parameters["price"].Replace('.', ',')),
                Currency = int.Parse(eventVm.Parameters["income"]),
            };

            context.Events.Add(e);
            context.CurrencyPurchases.Add(purchase);
        }

        private void SaveItemPurchase(EventViewModel eventVm)
        {
            var e = GetNewEvent(eventVm);
            var purchase = new ItemPurchase
            {
                Id = e.Id,
                ItemName = eventVm.Parameters["item"],
                Price = int.Parse(eventVm.Parameters["price"]),
            };

            context.Events.Add(e);
            context.ItemPurchases.Add(purchase);
        }

        private void SaveStageStart(EventViewModel eventVm)
        {
            var e = GetNewEvent(eventVm);
            var purchase = new StageStart
            {
                Id = e.Id,
                Stage = int.Parse(eventVm.Parameters["stage"]),
            };

            context.Events.Add(e);
            context.StageStarts.Add(purchase);
        }

        private void SaveStageEnd(EventViewModel eventVm)
        {
            var e = GetNewEvent(eventVm);
            var purchase = new StageEnd
            {
                Id = e.Id,
                Stage = int.Parse(eventVm.Parameters["stage"]),
                Win = bool.Parse(eventVm.Parameters["win"]),
                Time = int.Parse(eventVm.Parameters["time"]),
                Currency = int.Parse(eventVm.Parameters["income"]),
            };

            context.Events.Add(e);
            context.StageEnds.Add(purchase);
        }

        private Event GetNewEvent(EventViewModel eventVm)
        {
            return new Event
            {
                Id = Guid.NewGuid(),
                Date = eventVm.Date,
                UserId = eventVm.Udid,
            };
        }
    }
}
