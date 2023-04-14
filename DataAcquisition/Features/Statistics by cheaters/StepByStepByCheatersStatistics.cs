using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_cheaters
{
    public static class StepByStepByCheatersStatistics
    {
        public static ExcelPackage AddStepByStepByCheatersStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Step-by-step by cheaters statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Step-by-step by cheaters statistics");

            worksheet.Cells["A1"].Value = "Stage";
            worksheet.Cells["A1:A2"].Merge = true;
            worksheet.Cells["B1"].Value = "Starts";
            worksheet.Cells["B1:C1"].Merge = true;
            worksheet.Cells["B2"].Value = "Cheaters";
            worksheet.Cells["C2"].Value = "Non cheaters";
            worksheet.Cells["D1"].Value = "Ends";
            worksheet.Cells["D1:E1"].Merge = true;
            worksheet.Cells["D2"].Value = "Cheaters";
            worksheet.Cells["E2"].Value = "Non cheaters";
            worksheet.Cells["F1"].Value = "Wins";
            worksheet.Cells["F1:G1"].Merge = true;
            worksheet.Cells["F2"].Value = "Cheaters";
            worksheet.Cells["G2"].Value = "Non cheaters";
            worksheet.Cells["H1"].Value = "Currency";
            worksheet.Cells["H1:I1"].Merge = true;
            worksheet.Cells["H2"].Value = "Cheaters";
            worksheet.Cells["I2"].Value = "Non cheaters";
            worksheet.Cells["J1"].Value = "USD";
            worksheet.Cells["J1:K1"].Merge = true;
            worksheet.Cells["J2"].Value = "Cheaters";
            worksheet.Cells["K2"].Value = "Non cheaters";

            var stages = context.StageStarts
                .GroupBy(stageStart => stageStart.Stage)
                .Select(group => new
                {
                    Stage = group.Key.Value, 
                    StartsCheaters = group.Count(x => x.IdNavigation.User.IsCheater.Equals(true)),
                    StartsNonCheaters = group.Count(x => x.IdNavigation.User.IsCheater.Equals(false))
                })
                .Join(
                    context.StageEnds
                        .GroupBy(stageEnd => stageEnd.Stage)
                        .Select(group => new
                        {
                            Stage = group.Key.Value,
                            EndsCheaters = group
                                .Count(x => x.IdNavigation.User.IsCheater.Equals(true)),
                            WinAmountCheaters = group
                                .Where(x => x.IdNavigation.User.IsCheater.Equals(true))
                                .Count(x => (bool)x.Win),
                            CurrencyCheaters = group
                                .Where(x => x.IdNavigation.User.IsCheater.Equals(true))
                                .Sum(x => (bool)x.Win ? x.Currency : 0),
                            USDCheaters = group
                                .Where(x => x.IdNavigation.User.IsCheater.Equals(true))
                                .Sum(x => (bool)x.Win ? x.Currency : 0) * Utilities.GetEventUSDRate(context),
                            
                            EndsNonCheaters = group
                                .Count(x => x.IdNavigation.User.IsCheater.Equals(false)),
                            WinAmountNonCheaters = group
                                .Where(x => x.IdNavigation.User.IsCheater.Equals(false))
                                .Count(x => (bool)x.Win),
                            CurrencyNonCheaters = group
                                .Where(x => x.IdNavigation.User.IsCheater.Equals(false))
                                .Sum(x => (bool)x.Win ? x.Currency : 0),
                            USDNonCheaters = group
                                .Where(x => x.IdNavigation.User.IsCheater.Equals(false))
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
                worksheet.Cells[String.Concat("B", i + 3)].Value = stages[i].stageStart.StartsCheaters;
                worksheet.Cells[String.Concat("C", i + 3)].Value = stages[i].stageStart.StartsNonCheaters;
                worksheet.Cells[String.Concat("D", i + 3)].Value = stages[i].stageEnd.EndsCheaters;
                worksheet.Cells[String.Concat("E", i + 3)].Value = stages[i].stageEnd.EndsNonCheaters;
                worksheet.Cells[String.Concat("F", i + 3)].Value = stages[i].stageEnd.WinAmountCheaters;
                worksheet.Cells[String.Concat("G", i + 3)].Value = stages[i].stageEnd.WinAmountNonCheaters;
                worksheet.Cells[String.Concat("H", i + 3)].Value = stages[i].stageEnd.CurrencyCheaters;
                worksheet.Cells[String.Concat("I", i + 3)].Value = stages[i].stageEnd.CurrencyNonCheaters;
                worksheet.Cells[String.Concat("J", i + 3)].Value = stages[i].stageEnd.USDCheaters;
                worksheet.Cells[String.Concat("K", i + 3)].Value = stages[i].stageEnd.USDNonCheaters;
            }

            Console.WriteLine("Step-by-step by gender statistics added");

            return excelPackage;
        }
    }
}