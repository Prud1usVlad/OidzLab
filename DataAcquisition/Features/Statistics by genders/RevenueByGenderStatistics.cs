using DataAcquisition.Models;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_genders
{
    public static class RevenueByGenderStatistics
    {
        public static ExcelPackage AddRevenueByGenderStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Revenue by gender statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Revenue by gender statistics");

            worksheet.Cells["A1"].Value = "Day";
            worksheet.Cells["B1"].Value = "Revenue male, $";
            worksheet.Cells["C1"].Value = "Revenue female, $";
            
            var data = context.Events
                .Where(e => e.Type == 6)
                .GroupBy(e => e.Date)
                .Select(group => new
                {
                    Date = group.Key,
                    RevenueMale = group
                        .Where(x=> x.User.Gender.Equals("male"))
                        .Sum(i => i.CurrencyPurchase.Price),
                    RevenueFemale = group
                        .Where(x=> x.User.Gender.Equals("female"))
                        .Sum(i => i.CurrencyPurchase.Price)
                })
                .OrderBy(x=>x.Date)
                .ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value =
                    DateOnly.FromDateTime(data[i].Date.Value).ToString();
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].RevenueMale;
                worksheet.Cells[String.Concat("C", i + 2)].Value = data[i].RevenueFemale;
            }
            
            Console.WriteLine("Revenue by gender statistics added");
            
            return excelPackage;
        }
    }
}