using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features
{
    public static class UniqueUsersByStages
    {
        public static ExcelPackage AddUniqueUsersByStagesStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Unique users by stages init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Unique users by stages");

            worksheet.Cells["A1"].Value = "Stage";
            worksheet.Cells["B1"].Value = "Unique user amount";
            worksheet.Cells["C1"].Value = "Stage starts";
            worksheet.Cells["C1"].Value = "Stage ends";
            worksheet.Cells["C1"].Value = "Stage wins";

            var data = context.StageStarts
                .GroupBy(x => x.Stage)
                .Select(x => new
                {
                    Stage = x.Key,
                    UniqueUserByStageAmount = x.GroupBy(y => y.IdNavigation.UserId).Count(),
                    StageStart = x.Count()
                })
                .Join(
                    context.StageEnds
                        .GroupBy(stageEnd => stageEnd.Stage)
                        .Select(x => new
                        {
                            Stage = x.Key.Value,
                            Ends = x.Count(),
                            WinAmount = x.Count(x => (bool)x.Win)
                        }),
                    stageStart => stageStart.Stage,
                    stageEnd => stageEnd.Stage,
                    (stageStart, stageEnd) => new { stageStart, stageEnd })
                .OrderBy(x=>x.stageStart.Stage)
                .ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value = data[i].stageStart.Stage;
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].stageStart.UniqueUserByStageAmount;
                worksheet.Cells[String.Concat("C", i + 2)].Value = data[i].stageStart.StageStart;
                worksheet.Cells[String.Concat("D", i + 2)].Value = data[i].stageEnd.Ends;
                worksheet.Cells[String.Concat("E", i + 2)].Value = data[i].stageEnd.WinAmount;
            }
            
            Console.WriteLine("Unique users by stages added");
            
            return excelPackage;
        }
    }
}