using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_countries
{
    public static class PreliminaryByCountriesStatistics
    {
        public static ExcelPackage AddPreliminaryByCountriesStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Preliminary by countries statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Preliminary by countries statistics");

            var countries = Utilities.GetCountries(context);
            var countryAmount = countries.Count;
            
            worksheet.Cells["A1"].Value = "Item name";
            worksheet.Cells["A1:A2"].Merge = true;
            worksheet.Cells["B1"].Value = "Item amount";
            worksheet.Cells[String.Concat(
                String.Concat(Utilities.GetCellColumnAddress(2 + countryAmount * 0), "1"),
                ":",
                String.Concat(Utilities.GetCellColumnAddress(1 + countryAmount * 1), "1")
            )].Merge = true;
            for (int i = 0; i < countryAmount; i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(i + 2 + countryAmount * 0), "2")]
                    .Value = countries[i];
            }
            
            worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(countryAmount * 1 + 2), "1")].Value = "Currency";
            worksheet.Cells[String.Concat(
                String.Concat(Utilities.GetCellColumnAddress(2 + countryAmount * 1), "1"),
                ":",
                String.Concat(Utilities.GetCellColumnAddress(1 + countryAmount * 2), "1")
            )].Merge = true;
            for (int i = 0; i < countryAmount; i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(i + 2 + countryAmount * 1), "2")]
                    .Value = countries[i];
            }

            worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(countryAmount * 2 + 2), "1")].Value = "USD";
            worksheet.Cells[String.Concat(
                String.Concat(Utilities.GetCellColumnAddress(2 + countryAmount * 2), "1"),
                ":",
                String.Concat(Utilities.GetCellColumnAddress(1 + countryAmount * 3), "1")
            )].Merge = true;
            for (int i = 0; i < countryAmount; i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(i + 2 + countryAmount * 2), "2")]
                    .Value = countries[i];
            }

            var items = context.ItemPurchases
                .GroupBy(purchase => purchase.ItemName)
                .Select(group =>
                    new
                    {
                        ItemName = group.Key,
                        Countries = group
                            .GroupBy(o => o.IdNavigation.User.Country)
                            .Select(x => new
                            {
                                Country = x.Key,
                                ItemAmount = x.Count(),
                                Currency = x.Sum(y => y.Price),
                                USD = x.Sum(y => y.Price) * Utilities.GetEventUSDRate(context)
                            })
                    })
                .OrderBy(x=>x.ItemName)
                .ToList();

            for (int i = 0; i < items.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value = items[i].ItemName;
                
                for (int j = 0; j < countryAmount; j++)
                {
                    worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(j + 2 + countryAmount * 0), (i + 3).ToString())]
                        .Value = 0;
                    worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(j + 2 + countryAmount * 1), (i + 3).ToString())]
                        .Value = 0;
                    worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(j + 2 + countryAmount * 2), (i + 3).ToString())]
                        .Value = 0;
                }
                
                foreach (var country in items[i].Countries)
                {
                    worksheet.Cells[String.Concat(
                            Utilities.GetCellColumnAddress(countries.IndexOf(country.Country) + 2 + countryAmount * 0), 
                            (i + 3).ToString())]
                        .Value = country.ItemAmount;
                }
                
                foreach (var country in items[i].Countries)
                {
                    worksheet.Cells[String.Concat(
                            Utilities.GetCellColumnAddress(countries.IndexOf(country.Country) + 2 + countryAmount * 1), 
                            (i + 3).ToString())]
                        .Value = country.Currency;
                }
                
                foreach (var country in items[i].Countries)
                {
                    worksheet.Cells[String.Concat(
                            Utilities.GetCellColumnAddress(countries.IndexOf(country.Country) + 2 + countryAmount * 2), 
                            (i + 3).ToString())]
                        .Value = country.USD;
                }
            }

            Console.WriteLine("Preliminary by countries statistics added");
            
            return excelPackage;
        }
    }
}