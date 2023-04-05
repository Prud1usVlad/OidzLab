﻿using DataAcquisition.Features;
using DataAcquisition.Models;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

namespace DataAcquisition
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // This will get the current PROJECT directory
            string projectDirectory = Directory.GetParent(
                Directory.GetCurrentDirectory()).Parent.Parent.FullName;

            DirectoryInfo srcDir = new DirectoryInfo(projectDirectory + "\\Src");
            DirectoryInfo resultsDir = new DirectoryInfo(projectDirectory + "\\Results");

            FileInfo[] files = resultsDir.GetFiles("*.json");
            int count = files.Length;
            
            // Variables for partial file processing
            var curFiles = new List<FileInfo>();
            int start = 240;
            int end = count;

            //To create new files, clear from user repetitions 
            //RemoveUserRepetitions(files, resultsDir);

            //To upload data from chosen files
            //UploadFilesToDatabase(files, start, end);
            
            //To create spreadsheet with metrics
            Console.WriteLine(DateTime.Now);
            GenarateMetricsSpreadsheet(resultsDir);
            Console.WriteLine(DateTime.Now);

            Console.ReadLine();
        }

        public static void ClearCurrentConsoleLine()
        {
            int currentLineCursor = Console.CursorTop;
            Console.SetCursorPosition(0, Console.CursorTop);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, currentLineCursor);
        }

        public static void RemoveUserRepetitions(FileInfo[] files, DirectoryInfo resultsDir)
        {
            // var c = new RepetitionsRemover();
            //
            // for (int i = 0; i < files.Length; i++)
            // {
            //     c.RemoveRepetitions(files[i].FullName, resultsDir.ToString(), i);
            // }
        }

        public static void UploadFilesToDatabase(FileInfo[] files, int startIndex, int endIndex)
        {
            var etl = new EtlCore();
            
            for (int i = startIndex; i < endIndex; i++)
            {
                var file = files[i];
                
                Console.WriteLine("Processing file: " + file.Name + " ...");
                etl.ReadData(file.FullName);
                Console.SetCursorPosition(0, Console.CursorTop - 1);
                ClearCurrentConsoleLine();
                Console.WriteLine("File processed!");
                Console.WriteLine((endIndex - i) + "/" + (endIndex - startIndex) + " Files to go!");
            }
        }

        public static void GenarateMetricsSpreadsheet(DirectoryInfo resultsDir)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var context = new OidzDbContext();
            context.ChangeTracker.AutoDetectChangesEnabled = false;
            context.Database.SetCommandTimeout((int)TimeSpan.FromMinutes(30).TotalSeconds);
            
            //Creating excel file and filling it
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //Filling file with sheets
                excelPackage
                    // .AddNewUsersStatisticsSheet(context)
                    // .AddDauStatisticsSheet(context)
                    // .AddMauStatisticsSheet(context)
                    // .AddRevenueStatisticsSheet(context)
                    // .AddCurrencyRateStatisticsSheet(context)
                    // .AddStepByStepStatisticsSheet(context)
                    // .AddPreliminaryStatisticsSheet(context)
                    .AddItemsPerDayStatisticsSheet(context);
                
                excelPackage.SaveAs(
                    new FileInfo(
                        String.Concat(resultsDir.ToString(), "\\Sheets.xlsx")));
            }
        } 
    }
}