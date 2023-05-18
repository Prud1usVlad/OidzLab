using DataAcquisition.Models;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_cheaters
{
    public static class MauByCheatersStatistics
    {
        public static ExcelPackage AddMauByCheatersStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("MAU by cheaters statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("MAU by cheaters statistics");

            worksheet.Cells["A1"].Value = "Month";
            worksheet.Cells["B1"].Value = "MAU cheaters";
            worksheet.Cells["C1"].Value = "MAU non cheaters";

            var data = context.Events
                .GroupBy(e => new DateOnly(e.Date.Value.Year, e.Date.Value.Month, 1))
                .Select(group => new
                {
                    Date = group.Key,
                    Cheaters = group
                        .GroupBy(o => o.UserId)
                        .Count(x => x.Any(y => y.User.IsCheater.Equals(true))),
                    NonCheaters = group
                        .GroupBy(o => o.UserId)
                        .Count(x => x.Any(y => y.User.IsCheater.Equals(false)))
                })
                .ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value = data[i].Date.ToString();
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].Cheaters;
                worksheet.Cells[String.Concat("C", i + 2)].Value = data[i].NonCheaters;
            }
            
            Console.WriteLine("Mau by cheaters statistics added");
            
            return excelPackage;
        }
    }
}