using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_countries
{
    public static class StepByStepByCountriesStatistics
    {
        public static ExcelPackage AddStepByStepByCountriesStatisticsSheet(this ExcelPackage excelPackage,
            OidzDbContext context)
        {
            Console.WriteLine("Step-by-step by countries statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Step-by-step by countries statistics");

            var countries = Utilities.GetCountries(context);
            var countryAmount = countries.Count;

            worksheet.Cells["A1"].Value = "Stage";
            worksheet.Cells["A1:A2"].Merge = true;
            worksheet.Cells["B1"].Value = "Starts";
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

            worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(countryAmount * 1 + 2), "1")].Value = "Ends";
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

            worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(countryAmount * 2 + 2), "1")].Value = "Wins";
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

            worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(countryAmount * 3 + 2), "1")].Value =
                "Currency";
            worksheet.Cells[String.Concat(
                String.Concat(Utilities.GetCellColumnAddress(2 + countryAmount * 3), "1"),
                ":",
                String.Concat(Utilities.GetCellColumnAddress(1 + countryAmount * 4), "1")
            )].Merge = true;
            for (int i = 0; i < countryAmount; i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(i + 2 + countryAmount * 3), "2")]
                    .Value = countries[i];
            }

            worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(countryAmount * 4 + 2), "1")].Value = "USD";
            worksheet.Cells[String.Concat(
                String.Concat(Utilities.GetCellColumnAddress(2 + countryAmount * 4), "1"),
                ":",
                String.Concat(Utilities.GetCellColumnAddress(1 + countryAmount * 5), "1")
            )].Merge = true;
            for (int i = 0; i < countryAmount; i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(i + 2 + countryAmount * 4), "2")]
                    .Value = countries[i];
            }

            var stages = context.StageStarts
                .GroupBy(stageStart => stageStart.Stage)
                .Select(group => new
                {
                    Stage = group.Key.Value, 
                    Countries = group
                        .GroupBy(o => o.IdNavigation.User.Country)
                        .Select(x => new
                        {
                            Country = x.Key,
                            Starts = x.Count()
                        })
                })
                .Join(
                    context.StageEnds
                        .GroupBy(stageEnd => stageEnd.Stage)
                        .Select(group => new
                        {
                            Stage = group.Key.Value,
                            Countries = group
                                .GroupBy(o => o.IdNavigation.User.Country)
                                .Select(x => new
                                {
                                    Country = x.Key,
                                    Ends = x.Count(),
                                    WinAmount = x.Count(y => (bool)y.Win),
                                    Currency = x.Sum(y => (bool)y.Win ? y.Currency : 0),
                                    USD = x.Sum(y => (bool)y.Win ? y.Currency : 0) * Utilities.GetEventUSDRate(context)
                                })
                        }),
                    stageStart => stageStart.Stage,
                    stageEnd => stageEnd.Stage,
                    (stageStart, stageEnd) => new { stageStart, stageEnd })
                .OrderBy(x => x.stageStart.Stage)
                .ToList();

            for (int i = 0; i < stages.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value = stages[i].stageStart.Stage;
                
                for (int j = 0; j < countryAmount; j++)
                {
                    worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(j + 2 + countryAmount * 0), (i + 3).ToString())]
                        .Value = 0;
                    worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(j + 2 + countryAmount * 1), (i + 3).ToString())]
                        .Value = 0;
                    worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(j + 2 + countryAmount * 2), (i + 3).ToString())]
                        .Value = 0;
                    worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(j + 2 + countryAmount * 3), (i + 3).ToString())]
                        .Value = 0;
                    worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(j + 2 + countryAmount * 4), (i + 3).ToString())]
                        .Value = 0;
                }
                
                foreach (var country in stages[i].stageStart.Countries)
                {
                    worksheet.Cells[String.Concat(
                            Utilities.GetCellColumnAddress(countries.IndexOf(country.Country) + 2 + countryAmount * 0), 
                            (i + 3).ToString())]
                        .Value = country.Starts;
                }
                
                foreach (var country in stages[i].stageEnd.Countries)
                {
                    worksheet.Cells[String.Concat(
                            Utilities.GetCellColumnAddress(countries.IndexOf(country.Country) + 2 + countryAmount * 1), 
                            (i + 3).ToString())]
                        .Value = country.Ends;
                }
                
                foreach (var country in stages[i].stageEnd.Countries)
                {
                    worksheet.Cells[String.Concat(
                            Utilities.GetCellColumnAddress(countries.IndexOf(country.Country) + 2 + countryAmount * 2), 
                            (i + 3).ToString())]
                        .Value = country.WinAmount;
                }
                
                foreach (var country in stages[i].stageEnd.Countries)
                {
                    worksheet.Cells[String.Concat(
                            Utilities.GetCellColumnAddress(countries.IndexOf(country.Country) + 2 + countryAmount * 3), 
                            (i + 3).ToString())]
                        .Value = country.Currency;
                }
                
                foreach (var country in stages[i].stageEnd.Countries)
                {
                    worksheet.Cells[String.Concat(
                            Utilities.GetCellColumnAddress(countries.IndexOf(country.Country) + 2 + countryAmount * 4), 
                            (i + 3).ToString())]
                        .Value = country.Ends;
                }
            }

            Console.WriteLine("Step-by-step by countries statistics added");

            return excelPackage;
        }
    }
}