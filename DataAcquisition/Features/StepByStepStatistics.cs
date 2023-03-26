using OfficeOpenXml;
using DataAcquisition.Models;
using DataAcquisition.Util;

namespace DataAcquisition.Features
{
    public static class StepByStepStatistics
    {
        public static ExcelPackage AddStepByStepStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Step-by-step statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Step-by-step statistics");

            worksheet.Cells["A1"].Value = "Stage";
            worksheet.Cells["B1"].Value = "Starts";
            worksheet.Cells["C1"].Value = "Ends";
            worksheet.Cells["D1"].Value = "Wins";
            worksheet.Cells["E1"].Value = "Currency";
            worksheet.Cells["F1"].Value = "USD";

            var stages = context.StageStarts
                .GroupBy(stageStart => stageStart.Stage)
                .Select(group => new { Stage = group.Key.Value, Starts = group.Count() })
                .Join(
                    context.StageEnds
                        .GroupBy(stageEnd => stageEnd.Stage)
                        .Select(group => new
                        {
                            Stage = group.Key.Value,
                            Ends = group.Count(),
                            WinAmount = group.Count(x => (bool)x.Win),
                            Currency = group.Sum(x => (bool)x.Win ? x.Currency : 0),
                            USD = group.Sum(x =>  (bool)x.Win ? x.Currency : 0) * Utilities.GetEventUSDRate(context)
                        }),
                    stageStart => stageStart.Stage,
                    stageEnd => stageEnd.Stage,
                    (stageStart, stageEnd) => new { stageStart, stageEnd })
                .OrderBy(x=>x.stageStart.Stage)
                .ToList();

            for (int i = 0; i < stages.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value = stages[i].stageStart.Stage;
                worksheet.Cells[String.Concat("B", i + 2)].Value = stages[i].stageStart.Starts;
                worksheet.Cells[String.Concat("C", i + 2)].Value = stages[i].stageEnd.Ends;
                worksheet.Cells[String.Concat("D", i + 2)].Value = stages[i].stageEnd.WinAmount;
                worksheet.Cells[String.Concat("E", i + 2)].Value = stages[i].stageEnd.Currency;
                worksheet.Cells[String.Concat("F", i + 2)].Value = stages[i].stageEnd.USD;
            }

            Console.WriteLine("Step-by-step statistics added");
            return excelPackage;
        }
    }
}