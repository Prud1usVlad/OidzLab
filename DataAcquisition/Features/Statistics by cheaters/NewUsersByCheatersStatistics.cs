using DataAcquisition.Models;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_cheaters
{
    public static class NewUsersByCheatersStatistics
    {
        public static ExcelPackage AddNewUsersByCheatersStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("New users by cheaters statistics initialized");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("New Users by cheaters");

            worksheet.Cells["A1"].Value = "Day";
            worksheet.Cells["B1"].Value = "Cheaters";
            worksheet.Cells["C1"].Value = "Non cheaters";
            worksheet.Cells["D1"].Value = "Cheaters total amount";
            worksheet.Cells["E1"].Value = "Non cheaters total amount";

            var data = context.Events
                .Where(i => i.Type == 2)
                .GroupBy(e => e.Date)
                .Select(group => new
                {
                    Date = group.Key,
                    Cheaters = group.Count(x => x.User.IsCheater.Equals(true)),
                    NonCheaters = group.Count(x => x.User.IsCheater.Equals(false))
                })
                .OrderBy(x=>x.Date)
                .ToList();

            
            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value =
                    DateOnly.FromDateTime(data[i].Date.Value).ToString();
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].Cheaters;
                worksheet.Cells[String.Concat("C", i + 2)].Value = data[i].NonCheaters;
            }
            
            worksheet.Cells[String.Concat("D", 2)].Value = data.Sum(x=>x.Cheaters);
            worksheet.Cells[String.Concat("E", 2)].Value = data.Sum(x=>x.NonCheaters);
            
            Console.WriteLine("New users by cheaters statistics added");
            
            return excelPackage;
        }
    }
}