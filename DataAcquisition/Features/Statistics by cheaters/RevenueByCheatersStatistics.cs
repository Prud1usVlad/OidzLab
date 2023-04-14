using DataAcquisition.Models;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_cheaters
{
    public static class RevenueByCheatersStatistics
    {
        public static ExcelPackage AddRevenueByCheatersStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Revenue by cheaters statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Revenue by cheaters statistics");

            worksheet.Cells["A1"].Value = "Day";
            worksheet.Cells["B1"].Value = "Revenue cheaters, $";
            worksheet.Cells["C1"].Value = "Revenue non cheaters, $";
            
            var data = context.Events
                .Where(e => e.Type == 6)
                .GroupBy(e => e.Date)
                .Select(group => new
                {
                    Date = group.Key,
                    RevenueCheaters = group
                        .Where(x=> x.User.IsCheater.Equals(true))
                        .Sum(i => i.CurrencyPurchase.Price),
                    RevenueNonCheaters = group
                        .Where(x=> x.User.IsCheater.Equals(false))
                        .Sum(i => i.CurrencyPurchase.Price)
                })
                .OrderBy(x=>x.Date)
                .ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value =
                    DateOnly.FromDateTime(data[i].Date.Value).ToString();
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].RevenueCheaters;
                worksheet.Cells[String.Concat("C", i + 2)].Value = data[i].RevenueNonCheaters;
            }
            
            Console.WriteLine("Revenue by cheaters statistics added");
            
            return excelPackage;
        }
    }
}