using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_genders
{
    public static class ItemsPerDayByCheatersStatistics
    {
        public static ExcelPackage AddItemsPerDayByGenderStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Items-per-day by gender statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Items-per-day by gender statistics");

            worksheet.Cells["A1"].Value = "Date";
            worksheet.Cells["A1:A2"].Merge = true;
            worksheet.Cells["B1"].Value = "Items amount";
            worksheet.Cells["B1:C1"].Merge = true;
            worksheet.Cells["B2"].Value = "Male";
            worksheet.Cells["C2"].Value = "Female";
            worksheet.Cells["D1"].Value = "USD";
            worksheet.Cells["D1:E1"].Merge = true;
            worksheet.Cells["D2"].Value = "Male";
            worksheet.Cells["E2"].Value = "Female";
            
            var items = context.ItemPurchases
                .GroupBy(purchase => purchase.IdNavigation.Date)
                .Select(group =>
                    new
                    {
                        Date = group.Key,
                        ItemAmountMale = group
                            .Count(x => x.IdNavigation.User.Gender.Equals("male")),
                        USDMale = group
                            .Where(x => x.IdNavigation.User.Gender.Equals("male"))
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),
                        
                        ItemAmountFemale = group
                            .Count(x => x.IdNavigation.User.Gender.Equals("female")),
                        USDFemale = group
                            .Where(x => x.IdNavigation.User.Gender.Equals("female"))
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context)
                    })
                .OrderBy(x=>x.Date)
                .ToList();

            for (int i = 0; i < items.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 3)].Value = DateOnly.FromDateTime(items[i].Date.Value).ToString();
                worksheet.Cells[String.Concat("B", i + 3)].Value = items[i].ItemAmountMale;
                worksheet.Cells[String.Concat("C", i + 3)].Value = items[i].ItemAmountFemale;
                worksheet.Cells[String.Concat("D", i + 3)].Value = items[i].USDMale;
                worksheet.Cells[String.Concat("E", i + 3)].Value = items[i].ItemAmountFemale;
            }

            Console.WriteLine("Items-per-day by gender statistics added");
            
            return excelPackage;
        }
    }
}