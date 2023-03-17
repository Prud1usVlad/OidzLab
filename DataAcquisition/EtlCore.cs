﻿using System;
using System.Collections.Generic;
using System.Diagnostics.Metrics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using DataAcquisition.Models;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;
using Newtonsoft.Json;



namespace DataAcquisition
{
    public class EtlCore
    {
        private List<EventViewModel> rawData;
        private OidzDbContext context;
        private List<Guid> cachedUsers;

        public EtlCore()
        {
            context = new OidzDbContext();
        }

        public void ReadData(string path)
        {
            using (StreamReader file = File.OpenText(path))
            {
                JsonSerializer serializer = new JsonSerializer();
                rawData = (List<EventViewModel>)serializer.Deserialize(file, typeof(List<EventViewModel>));
            }

            using (var ts = CreateTransactionScope(TimeSpan.FromMinutes(60)))
            {
                context = null;
                try
                {
                    context = new OidzDbContext();
                    context.ChangeTracker.AutoDetectChangesEnabled = false;
                    cachedUsers = context.Users.Select(x => x.Id).ToList();

                    int count = 0;
                    foreach (var piece in rawData)
                    {
                        context = SaveData(context, piece, count, 1000, true);
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

        private OidzDbContext SaveData(
            OidzDbContext context,
            EventViewModel entity, 
            int count, 
            int commitCount, 
            bool recreateContext)
        {
            switch (entity.Event_id)
            {
                case 1:
                    SaveLaunch(entity);
                    break;
                case 2:
                    SaveFirstLaunch(entity);
                    break;
                case 3:
                    SaveStageStart(entity);
                    break;
                case 4:
                    SaveStageEnd(entity);
                    break;
                case 5:
                    SaveItemPurchase(entity);
                    break;
                case 6:
                    SaveCurrencyPurchase(entity);
                    break;
            }

            if (count % commitCount == 0)
            {
                context.SaveChanges();
                if (count != 0)
                {
                    Console.SetCursorPosition(0, Console.CursorTop - 1);
                    Program.ClearCurrentConsoleLine();
                }
                Console.WriteLine(count + " pieces processed. " + (rawData.Count() - count) + " to go...");


                if (recreateContext)
                {
                    context.Dispose();
                    context = new OidzDbContext();
                    context.ChangeTracker.AutoDetectChangesEnabled = false;
                }
            }

            return context;
        }

        private void SaveLaunch(EventViewModel eventVm)
        {
            var e = GetNewEvent(eventVm);

            context.Events.Add(e);
        }

        private void SaveFirstLaunch(EventViewModel eventVm)
        {

            if (!cachedUsers.Contains(eventVm.Udid))
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
                cachedUsers.Add(user.Id);
                context.SaveChanges();
            }

            //User user = null;
            //Event e = null;

            //// To avoid duplicate values
            //try 
            //{
            //    var e = GetNewEvent(eventVm);
            //    user = new User
            //    {
            //        Id = eventVm.Udid,
            //        Gender = eventVm.Parameters["gender"],
            //        Age = int.Parse(eventVm.Parameters["age"]),
            //        Country = eventVm.Parameters["country"],
            //    };

            //    context.Events.Add(e);
            //    context.Users.Add(user);
            //    context.SaveChanges();
            //}
            //catch (Exception ex) 
            //{
            //    context.Users.Local.Remove(user);
            //    context.Events.Local.Remove()
            //    var c1 = context.Users.Local;
            //    var c2 = context.Events.Local;
            //}
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
                Type = eventVm.Event_id
            };
        }

        private void SetTransactionManagerField(string fieldName, object value)
        {
            typeof(TransactionManager).GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Static).SetValue(null, value);
        }

        public TransactionScope CreateTransactionScope(TimeSpan timeout)
        {
            // or for netcore / .net5+ use these names instead:
            //    s_cachedMaxTimeout
            //    s_maximumTimeout
            SetTransactionManagerField("s_cachedMaxTimeout", true);
            SetTransactionManagerField("s_maximumTimeout", timeout);
            return new TransactionScope(TransactionScopeOption.RequiresNew, timeout);
        }
    }
}
