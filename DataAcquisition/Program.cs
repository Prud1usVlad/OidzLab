using DataAcquisition.Models;
using System.IO;
using System.Security.Cryptography;

namespace DataAcquisition
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var context = new PostgresContext();
            var etl = new EtlCore();


            DirectoryInfo dir = new DirectoryInfo("D:\\NURE\\ThirdCourse\\SecondSemester\\oidz\\labs\\DataAcquisition\\DataAcquisition\\Src\\"); //Assuming Test is your Folder
            FileInfo[] files = dir.GetFiles("*.json");
            int count = files.Length;

            foreach (FileInfo file in files)
            {
                Console.WriteLine("Processing file: " + file.Name + " ...");
                etl.ReadData(file.FullName);
                Console.WriteLine("File processed!");
                Console.WriteLine(--count + "/" + files.Length + " Files to go!");
                break;
            }



            Console.ReadLine();
        }
    }
}