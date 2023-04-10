using DataAcquisition.Models;
using OfficeOpenXml;

namespace DataAcquisition.Features
{
    public static partial class CurrencyMetrics
    {
        public static ExcelPackage AddRevenueStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Revenue statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Revenue statistics");

            worksheet.Cells["A1"].Value = "Day";
            worksheet.Cells["B1"].Value = "Revenue, $";

            var data = context.Events
                .Where(e => e.Type == 6)
                .GroupBy(e => e.Date)
                .Select(group => new
                {
                    Date = group.Key,
                    Revenue = group.Sum(i => i.CurrencyPurchase.Price),
                })
                .OrderBy(x=>x.Date)
                .ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value =
                    DateOnly.FromDateTime(data[i].Date.Value).ToString();
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].Revenue;
            }
            
            Console.WriteLine("Revenue statistics added");
            
            return excelPackage;
        }
    }
}
