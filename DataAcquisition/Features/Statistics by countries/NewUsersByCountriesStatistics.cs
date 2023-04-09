using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_countries
{
    public static class NewUsersByCountriesStatistics
    {
        public static ExcelPackage AddNewUsersByCountriesStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("New users by countries statistics initialized");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("New Users by countries");

            var countries = Utilities.GetCountries(context);
            var countryAmount = countries.Count;
            
            worksheet.Cells["A1"].Value = "Day";
           
            for (int i = 0; i < countryAmount; i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(i + 2), "1")]
                    .Value = countries[i];
            }

            var data = context.Events
                .Where(i => i.Type == 2)
                .GroupBy(e => e.Date)
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
                            (i + 3).ToString())]
                        .Value = country.Count;
                }
            }
            
            Console.WriteLine("New users by countries statistics added");
            
            return excelPackage;
        }
    }
}