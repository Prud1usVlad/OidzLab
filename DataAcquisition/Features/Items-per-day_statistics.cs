using OfficeOpenXml;
using DataAcquisition.Models;
using DataAcquisition.Util;

namespace DataAcquisition.Features
{
    public static class ItemsPerDayStatistics
    {
        public static ExcelPackage AddItemsPerDayStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Items-per-day statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Items-per-day statistics");

            worksheet.Cells["A1"].Value = "Date";
            worksheet.Cells["B1"].Value = "Items amount";
            worksheet.Cells["C1"].Value = "USD";
            
            var items = context.ItemPurchases
                .GroupBy(purchase => purchase.IdNavigation.Date)
                .Select(group =>
                    new
                    {
                        Date = group.Key,
                        ItemAmount = group.Count(),
                        USD = group.Sum(x => x.Price) * Utilities.GetEventUSDRate(context)
                    })
                .OrderBy(x=>x.Date)
                .ToList();

            for (int i = 0; i < items.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value = DateOnly.FromDateTime(items[i].Date.Value).ToString();
                worksheet.Cells[String.Concat("B", i + 2)].Value = items[i].ItemAmount;
                worksheet.Cells[String.Concat("C", i + 2)].Value = items[i].USD;
            }

            Console.WriteLine("Items-per-day statistics added");
            return excelPackage;
        }
    }
}