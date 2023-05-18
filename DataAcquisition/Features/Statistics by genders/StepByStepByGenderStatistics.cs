using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_genders
{
    public static class StepByStepByCheatersStatistics
    {
        public static ExcelPackage AddStepByStepByGenderStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Step-by-step by gender statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Step-by-step by gender statistics");

            worksheet.Cells["A1"].Value = "Stage";
            worksheet.Cells["A1:A2"].Merge = true;
            worksheet.Cells["B1"].Value = "Starts";
            worksheet.Cells["B1:C1"].Merge = true;
            worksheet.Cells["B2"].Value = "Male";
            worksheet.Cells["C2"].Value = "Female";
            worksheet.Cells["D1"].Value = "Ends";
            worksheet.Cells["D1:E1"].Merge = true;
            worksheet.Cells["D2"].Value = "Male";
            worksheet.Cells["E2"].Value = "Female";
            worksheet.Cells["F1"].Value = "Wins";
            worksheet.Cells["F1:G1"].Merge = true;
            worksheet.Cells["F2"].Value = "Male";
            worksheet.Cells["G2"].Value = "Female";
            worksheet.Cells["H1"].Value = "Currency";
            worksheet.Cells["H1:I1"].Merge = true;
            worksheet.Cells["H2"].Value = "Male";
            worksheet.Cells["I2"].Value = "Female";
            worksheet.Cells["J1"].Value = "USD";
            worksheet.Cells["J1:K1"].Merge = true;
            worksheet.Cells["J2"].Value = "Male";
            worksheet.Cells["K2"].Value = "Female";

            var stages = context.StageStarts
                .GroupBy(stageStart => stageStart.Stage)
                .Select(group => new
                {
                    Stage = group.Key.Value, 
                    StartsMale = group.Count(x => x.IdNavigation.User.Gender.Equals("male")),
                    StartsFemale = group.Count(x => x.IdNavigation.User.Gender.Equals("female"))
                })
                .Join(
                    context.StageEnds
                        .GroupBy(stageEnd => stageEnd.Stage)
                        .Select(group => new
                        {
                            Stage = group.Key.Value,
                            EndsMale = group
                                .Count(x => x.IdNavigation.User.Gender.Equals("male")),
                            WinAmountMale = group
                                .Where(x => x.IdNavigation.User.Gender.Equals("male"))
                                .Count(x => (bool)x.Win),
                            CurrencyMale = group
                                .Where(x => x.IdNavigation.User.Gender.Equals("male"))
                                .Sum(x => (bool)x.Win ? x.Currency : 0),
                            USDMale = group
                                .Where(x => x.IdNavigation.User.Gender.Equals("male"))
                                .Sum(x => (bool)x.Win ? x.Currency : 0) * Utilities.GetEventUSDRate(context),
                            
                            EndsFemale = group
                                .Count(x => x.IdNavigation.User.Gender.Equals("female")),
                            WinAmountFemale = group
                                .Where(x => x.IdNavigation.User.Gender.Equals("female"))
                                .Count(x => (bool)x.Win),
                            CurrencyFemale = group
                                .Where(x => x.IdNavigation.User.Gender.Equals("female"))
                                .Sum(x => (bool)x.Win ? x.Currency : 0),
                            USDFemale = group
                                .Where(x => x.IdNavigation.User.Gender.Equals("female"))
                                .Sum(x => (bool)x.Win ? x.Currency : 0) * Utilities.GetEventUSDRate(context)
                        }),
                    stageStart => stageStart.Stage,
                    stageEnd => stageEnd.Stage,
                    (stageStart, stageEnd) => new { stageStart, stageEnd })
                .OrderBy(x => x.stageStart.Stage)
                .ToList();

            for (int i = 0; i < stages.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 3)].Value = stages[i].stageStart.Stage;
                worksheet.Cells[String.Concat("B", i + 3)].Value = stages[i].stageStart.StartsMale;
                worksheet.Cells[String.Concat("C", i + 3)].Value = stages[i].stageStart.StartsFemale;
                worksheet.Cells[String.Concat("D", i + 3)].Value = stages[i].stageEnd.EndsMale;
                worksheet.Cells[String.Concat("E", i + 3)].Value = stages[i].stageEnd.EndsFemale;
                worksheet.Cells[String.Concat("F", i + 3)].Value = stages[i].stageEnd.WinAmountMale;
                worksheet.Cells[String.Concat("G", i + 3)].Value = stages[i].stageEnd.WinAmountFemale;
                worksheet.Cells[String.Concat("H", i + 3)].Value = stages[i].stageEnd.CurrencyMale;
                worksheet.Cells[String.Concat("I", i + 3)].Value = stages[i].stageEnd.CurrencyFemale;
                worksheet.Cells[String.Concat("J", i + 3)].Value = stages[i].stageEnd.USDMale;
                worksheet.Cells[String.Concat("K", i + 3)].Value = stages[i].stageEnd.USDFemale;
            }

            Console.WriteLine("Step-by-step by gender statistics added");

            return excelPackage;
        }
    }
}