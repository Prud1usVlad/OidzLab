using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_genders
{
    public static class PreliminaryByGenderStatistics
    {
        public static ExcelPackage AddPreliminaryByGenderStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Preliminary by gender statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Preliminary by gender statistics");

            worksheet.Cells["A1"].Value = "Item name";
            worksheet.Cells["A1:A2"].Merge = true;
            worksheet.Cells["B1"].Value = "Item amount";
            worksheet.Cells["B1:C1"].Merge = true;
            worksheet.Cells["B2"].Value = "Male";
            worksheet.Cells["C2"].Value = "Female";
            worksheet.Cells["D1"].Value = "Currency";
            worksheet.Cells["D1:E1"].Merge = true;
            worksheet.Cells["D2"].Value = "Male";
            worksheet.Cells["E2"].Value = "Female";
            worksheet.Cells["F1"].Value = "USD";
            worksheet.Cells["F1:G1"].Merge = true;
            worksheet.Cells["F2"].Value = "Male";
            worksheet.Cells["G2"].Value = "Female";
            
            var items = context.ItemPurchases
                .GroupBy(purchase => purchase.ItemName)
                .Select(group =>
                    new
                    {
                        ItemName = group.Key,
                        ItemAmountMale = group
                            .Count(x => x.IdNavigation.User.Gender.Equals("male")),
                        CurrencyMale = group
                            .Where(x => x.IdNavigation.User.Gender.Equals("male"))
                            .Sum(x => x.Price),
                        USDMale = group
                            .Where(x => x.IdNavigation.User.Gender.Equals("male"))
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),
                        
                        ItemAmountFemale = group
                            .Count(x => x.IdNavigation.User.Gender.Equals("female")),
                        CurrencyFemale = group
                            .Where(x => x.IdNavigation.User.Gender.Equals("female"))
                            .Sum(x => x.Price),
                        USDFemale = group
                            .Where(x => x.IdNavigation.User.Gender.Equals("female"))
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context)
                    })
                .OrderBy(x=>x.ItemName)
                .ToList();

            for (int i = 0; i < items.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 3)].Value = items[i].ItemName;
                worksheet.Cells[String.Concat("B", i + 3)].Value = items[i].ItemAmountMale;
                worksheet.Cells[String.Concat("C", i + 3)].Value = items[i].ItemAmountFemale;
                worksheet.Cells[String.Concat("D", i + 3)].Value = items[i].CurrencyMale;
                worksheet.Cells[String.Concat("E", i + 3)].Value = items[i].CurrencyFemale;
                worksheet.Cells[String.Concat("F", i + 3)].Value = items[i].USDMale;
                worksheet.Cells[String.Concat("G", i + 3)].Value = items[i].USDFemale;
            }

            Console.WriteLine("Preliminary by gender statistics added");
            
            return excelPackage;
        }
    }
}