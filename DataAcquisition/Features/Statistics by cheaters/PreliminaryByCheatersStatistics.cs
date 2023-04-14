using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_cheaters
{
    public static class PreliminaryByCheatersStatistics
    {
        public static ExcelPackage AddPreliminaryByCheatersStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Preliminary by cheaters statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Preliminary by cheaters statistics");

            worksheet.Cells["A1"].Value = "Item name";
            worksheet.Cells["A1:A2"].Merge = true;
            worksheet.Cells["B1"].Value = "Item amount";
            worksheet.Cells["B1:C1"].Merge = true;
            worksheet.Cells["B2"].Value = "Cheaters";
            worksheet.Cells["C2"].Value = "Non cheaters";
            worksheet.Cells["D1"].Value = "Currency";
            worksheet.Cells["D1:E1"].Merge = true;
            worksheet.Cells["D2"].Value = "Cheaters";
            worksheet.Cells["E2"].Value = "Non cheaters";
            worksheet.Cells["F1"].Value = "USD";
            worksheet.Cells["F1:G1"].Merge = true;
            worksheet.Cells["F2"].Value = "Cheaters";
            worksheet.Cells["G2"].Value = "Non cheaters";
            
            var items = context.ItemPurchases
                .GroupBy(purchase => purchase.ItemName)
                .Select(group =>
                    new
                    {
                        ItemName = group.Key,
                        ItemAmountCheaters = group
                            .Count(x => x.IdNavigation.User.IsCheater.Equals(true)),
                        CurrencyCheaters = group
                            .Where(x => x.IdNavigation.User.IsCheater.Equals(true))
                            .Sum(x => x.Price),
                        USDCheaters = group
                            .Where(x => x.IdNavigation.User.IsCheater.Equals(true))
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),
                        
                        ItemAmountNonCheaters = group
                            .Count(x => x.IdNavigation.User.IsCheater.Equals(false)),
                        CurrencyNonCheaters = group
                            .Where(x => x.IdNavigation.User.IsCheater.Equals(false))
                            .Sum(x => x.Price),
                        USDNonCheaters = group
                            .Where(x => x.IdNavigation.User.IsCheater.Equals(false))
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context)
                    })
                .OrderBy(x=>x.ItemName)
                .ToList();

            for (int i = 0; i < items.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 3)].Value = items[i].ItemName;
                worksheet.Cells[String.Concat("B", i + 3)].Value = items[i].ItemAmountCheaters;
                worksheet.Cells[String.Concat("C", i + 3)].Value = items[i].ItemAmountNonCheaters;
                worksheet.Cells[String.Concat("D", i + 3)].Value = items[i].CurrencyCheaters;
                worksheet.Cells[String.Concat("E", i + 3)].Value = items[i].CurrencyNonCheaters;
                worksheet.Cells[String.Concat("F", i + 3)].Value = items[i].USDCheaters;
                worksheet.Cells[String.Concat("G", i + 3)].Value = items[i].USDNonCheaters;
            }

            Console.WriteLine("Preliminary by cheaters statistics added");
            
            return excelPackage;
        }
    }
}