using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_cheaters
{
    public static class ItemsPerDayByCheatersStatistics
    {
        public static ExcelPackage AddItemsPerDayByCheatersStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Items-per-day by cheaters statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Items-per-day by cheaters statistics");

            worksheet.Cells["A1"].Value = "Date";
            worksheet.Cells["A1:A2"].Merge = true;
            worksheet.Cells["B1"].Value = "Items amount";
            worksheet.Cells["B1:C1"].Merge = true;
            worksheet.Cells["B2"].Value = "Cheaters";
            worksheet.Cells["C2"].Value = "Non cheaters";
            worksheet.Cells["D1"].Value = "USD";
            worksheet.Cells["D1:E1"].Merge = true;
            worksheet.Cells["D2"].Value = "Cheaters";
            worksheet.Cells["E2"].Value = "Non cheaters";
            
            var items = context.ItemPurchases
                .GroupBy(purchase => purchase.IdNavigation.Date)
                .Select(group =>
                    new
                    {
                        Date = group.Key,
                        ItemAmountCheaters = group
                            .Count(x => x.IdNavigation.User.IsCheater.Equals(true)),
                        USDCheaters = group
                            .Where(x => x.IdNavigation.User.IsCheater.Equals(true))
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),

                        ItemAmountNonCheaters = group
                            .Count(x => x.IdNavigation.User.IsCheater.Equals(false)),
                        USDNonCheaters = group
                            .Where(x => x.IdNavigation.User.IsCheater.Equals(false))
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context)
                    })
                .OrderBy(x=>x.Date)
                .ToList();

            for (int i = 0; i < items.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 3)].Value = DateOnly.FromDateTime(items[i].Date.Value).ToString();
                worksheet.Cells[String.Concat("B", i + 3)].Value = items[i].ItemAmountCheaters;
                worksheet.Cells[String.Concat("C", i + 3)].Value = items[i].ItemAmountNonCheaters;
                worksheet.Cells[String.Concat("D", i + 3)].Value = items[i].USDCheaters;
                worksheet.Cells[String.Concat("E", i + 3)].Value = items[i].USDNonCheaters;
            }

            Console.WriteLine("Items-per-day by cheaters statistics added");
            
            return excelPackage;
        }
    }
}