using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_countries
{
    public static class RevenueByCountriesStatistics
    {
        public static ExcelPackage AddRevenueByCountriesStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Revenue by countries statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Revenue by countries statistics");

            var countries = Utilities.GetCountries(context);
            var countryAmount = countries.Count;
            
            worksheet.Cells["A1"].Value = "Day";
            
            for (int i = 0; i < countryAmount; i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(i + 2), "1")]
                    .Value = countries[i];
            }

            var data = context.Events
                .Where(e => e.Type == 6)
                .GroupBy(e => e.Date)
                .Select(group => new
                {
                    Date = group.Key,
                    Countries = group
                        .GroupBy(o => o.User.Country)
                        .Select(x => new
                        {
                            Country = x.Key,
                            Revenue = x.Sum(i => i.CurrencyPurchase.Price)
                        })
                })
                .OrderBy(x=>x.Date)
                .ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value =
                    DateOnly.FromDateTime(data[i].Date.Value).ToString();
                
                for (int j = 0; j < countryAmount; j++)
                {
                    worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(j+2), (i + 3).ToString())]
                        .Value = 0;
                }
                
                foreach (var country in data[i].Countries)
                {
                    worksheet.Cells[String.Concat(
                            Utilities.GetCellColumnAddress(countries.IndexOf(country.Country)+2), 
                            (i + 2).ToString())]
                        .Value = country.Revenue;
                }
            }
            
            Console.WriteLine("Revenue by countries statistics added");
            
            return excelPackage;
        }
    }
}