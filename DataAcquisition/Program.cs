using DataAcquisition.Features;
using DataAcquisition.Models;
using OfficeOpenXml;

namespace DataAcquisition
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var context = new PostgresContext();
            var etl = new EtlCore();


            // This will get the current PROJECT directory
            string projectDirectory = Directory.GetParent(
                Directory.GetCurrentDirectory()).Parent.Parent.FullName;

            DirectoryInfo srcDir = new DirectoryInfo(projectDirectory + "\\Src");
            DirectoryInfo resultsDir = new DirectoryInfo(projectDirectory + "\\Results");

            FileInfo[] files = srcDir.GetFiles("*.json");
            int count = files.Length;

            //foreach ( FileInfo file in files )
            //{
            //    Console.WriteLine("Processing file: " + file.Name + " ...");
            //    etl.ReadData(file.FullName);
            //    Console.WriteLine("File processed!");
            //    Console.WriteLine(--count + "/" + files.Length + " Files to go!");
            //    break;
            //}


            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //Creating excel file and filling it
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //Filling file with sheets
                excelPackage
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
    }
}