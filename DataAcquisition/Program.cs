﻿using DataAcquisition.Features;
using DataAcquisition.Models;
using DataAcquisition.Util;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

namespace DataAcquisition
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var context = new OidzDbContext();
            context.ChangeTracker.AutoDetectChangesEnabled = false;
            context.Database.SetCommandTimeout((int)TimeSpan.FromMinutes(30).TotalSeconds);
            
            var etl = new EtlCore();


            // This will get the current PROJECT directory
            string projectDirectory = Directory.GetParent(
                Directory.GetCurrentDirectory()).Parent.Parent.FullName;

            DirectoryInfo srcDir = new DirectoryInfo(projectDirectory + "\\Src");
            DirectoryInfo resultsDir = new DirectoryInfo(projectDirectory + "\\Results");

            FileInfo[] files = resultsDir.GetFiles("*.json");
            int count = files.Length;

            var c = new RepetitionsRemover();

            var curFiles = new List<FileInfo>();
            int start = 20;
            int end = 25;

            ////Removing repetitions of user data
            //for (int i = 0; i < count; i++)
            //{
            //    c.RemoveRepetitions(files[i].FullName, resultsDir.ToString(), i);
            //}


            // for (int i = start; i < end; i++)
            // {
            //     curFiles.Add(files[i]);
            // }
            //
            // foreach (FileInfo file in curFiles)
            // {
            //     Console.WriteLine("Processing file: " + file.Name + " ...");
            //     etl.ReadData(file.FullName);
            //     context.SaveChanges();
            //     Console.SetCursorPosition(0, Console.CursorTop - 1);
            //     ClearCurrentConsoleLine();
            //     Console.WriteLine("File processed!");
            //     Console.WriteLine(--count + "/" + files.Length + " Files to go!");
            // }
            
            Console.WriteLine(DateTime.Now);
            Console.WriteLine(DateTime.Now.Date);
            Console.WriteLine(DateTime.Now.ToString());

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //Creating excel file and filling it
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //Filling file with sheets
                excelPackage
                    .AddNewUsersStatisticsSheet(context)
                    .AddDauStatisticsSheet(context)
                    .AddMauStatisticsSheet(context)
                    .AddRevenueStatisticsSheet(context)
                    .AddCurrencyRateStatisticsSheet(context)
                    .AddStepByStepStatisticsSheet(context)
                    .AddPreliminaryStatisticsSheet(context);


                excelPackage.SaveAs(
                    new FileInfo(
                    String.Concat(resultsDir.ToString(), "\\Sheets.xlsx")));
            }

            Console.WriteLine(resultsDir.ToString());
            Console.WriteLine(srcDir.ToString());
            Console.ReadLine();
        }

        public static void ClearCurrentConsoleLine()
        {
            int currentLineCursor = Console.CursorTop;
            Console.SetCursorPosition(0, Console.CursorTop);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, currentLineCursor);
        }
    }
}