using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_countries
{
    public static class MauByCountriesStatistics
    {
        public static ExcelPackage AddMauByCountriesStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("MAU by countries statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("MAU by countries statistics");

            var countries = Utilities.GetCountries(context);
            var countryAmount = countries.Count;
            
            worksheet.Cells["A1"].Value = "Month";
            
            for (int i = 0; i < countryAmount; i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(i + 2), "1")]
                    .Value = countries[i];
            }

            var data = context.Events
                .GroupBy(e => new DateOnly(e.Date.Value.Year, e.Date.Value.Month, 1))
                .Select(group => new
                {
                    Date = group.Key,
                    Countries = group
                        .GroupBy(o => o.User.Country)
                        .Select(x => new
                        {
                            Country = x.Key,
                            Count = x.GroupBy(y=> y.UserId).Count()
                        })
                })
                .ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value = data[i].Date.ToString();
                
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
                        .Value = country.Count;
                }
            }
            
            Console.WriteLine("Mau by countries statistics added");
            
            return excelPackage;
        }
    }
}