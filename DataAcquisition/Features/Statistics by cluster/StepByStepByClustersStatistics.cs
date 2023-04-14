using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_clusters
{
    public static class StepByStepByClustersStatistics
    {
        public static ExcelPackage AddStepByStepByClustersStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Step-by-step by clusters statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Step-by-step by clusters statistics");

            worksheet.Cells["A1"].Value = "Stage";
            worksheet.Cells["A1:A2"].Merge = true;

            worksheet.Cells["B1"].Value = "Starts";
            worksheet.Cells["B1:E1"].Merge = true;
            worksheet.Cells["B2"].Value = "Cluster O";
            worksheet.Cells["C2"].Value = "Cluster I";
            worksheet.Cells["D2"].Value = "Cluster II";
            worksheet.Cells["E2"].Value = "Cluster III";

            worksheet.Cells["F1"].Value = "Ends";
            worksheet.Cells["F1:I1"].Merge = true;
            worksheet.Cells["F2"].Value = "Cluster O";
            worksheet.Cells["G2"].Value = "Cluster I";
            worksheet.Cells["H2"].Value = "Cluster II";
            worksheet.Cells["I2"].Value = "Cluster III";

            worksheet.Cells["J1"].Value = "Wins";
            worksheet.Cells["J1:M1"].Merge = true;
            worksheet.Cells["J2"].Value = "Cluster O";
            worksheet.Cells["K2"].Value = "Cluster I";
            worksheet.Cells["L2"].Value = "Cluster II";
            worksheet.Cells["M2"].Value = "Cluster III";

            worksheet.Cells["N1"].Value = "Currency";
            worksheet.Cells["N1:Q1"].Merge = true;
            worksheet.Cells["N2"].Value = "Cluster O";
            worksheet.Cells["O2"].Value = "Cluster I";
            worksheet.Cells["P2"].Value = "Cluster II";
            worksheet.Cells["Q2"].Value = "Cluster III";

            worksheet.Cells["R1"].Value = "USD";
            worksheet.Cells["R1:U1"].Merge = true;
            worksheet.Cells["R2"].Value = "Cluster O";
            worksheet.Cells["S2"].Value = "Cluster I";
            worksheet.Cells["T2"].Value = "Cluster II";
            worksheet.Cells["U2"].Value = "Cluster III";

            var stages = context.StageStarts
                .GroupBy(stageStart => stageStart.Stage)
                .Select(group => new
                {
                    Stage = group.Key.Value, 
                    StartsClusterO = group.Count(x => x.IdNavigation.User.Cluster.Equals(0)),
                    StartsClusterI = group.Count(x => x.IdNavigation.User.Cluster.Equals(1)),
                    StartsClusterII = group.Count(x => x.IdNavigation.User.Cluster.Equals(2)),
                    StartsClusterIII = group.Count(x => x.IdNavigation.User.Cluster.Equals(3)),
                })
                .Join(
                    context.StageEnds
                        .GroupBy(stageEnd => stageEnd.Stage)
                        .Select(group => new
                        {
                            Stage = group.Key.Value,
                            EndsClusterO = group
                                .Count(x => x.IdNavigation.User.Cluster.Equals(0)),
                            WinAmountClusterO = group
                                .Where(x => x.IdNavigation.User.Cluster.Equals(0))
                                .Count(x => (bool)x.Win),
                            CurrencyClusterO = group
                                .Where(x => x.IdNavigation.User.Cluster.Equals(0))
                                .Sum(x => (bool)x.Win ? x.Currency : 0),
                            USDClusterO = group
                                .Where(x => x.IdNavigation.User.Cluster.Equals(0))
                                .Sum(x => (bool)x.Win ? x.Currency : 0) * Utilities.GetEventUSDRate(context),

                            EndsClusterI = group
                                .Count(x => x.IdNavigation.User.Cluster.Equals(1)),
                            WinAmountClusterI = group
                                .Where(x => x.IdNavigation.User.Cluster.Equals(1))
                                .Count(x => (bool)x.Win),
                            CurrencyClusterI = group
                                .Where(x => x.IdNavigation.User.Cluster.Equals(1))
                                .Sum(x => (bool)x.Win ? x.Currency : 0),
                            USDClusterI = group
                                .Where(x => x.IdNavigation.User.Cluster.Equals(1))
                                .Sum(x => (bool)x.Win ? x.Currency : 0) * Utilities.GetEventUSDRate(context),

                            EndsClusterII = group
                                .Count(x => x.IdNavigation.User.Cluster.Equals(2)),
                            WinAmountClusterII = group
                                .Where(x => x.IdNavigation.User.Cluster.Equals(2))
                                .Count(x => (bool)x.Win),
                            CurrencyClusterII = group
                                .Where(x => x.IdNavigation.User.Cluster.Equals(2))
                                .Sum(x => (bool)x.Win ? x.Currency : 0),
                            USDClusterII = group
                                .Where(x => x.IdNavigation.User.Cluster.Equals(2))
                                .Sum(x => (bool)x.Win ? x.Currency : 0) * Utilities.GetEventUSDRate(context),

                            EndsClusterIII = group
                                .Count(x => x.IdNavigation.User.Cluster.Equals(3)),
                            WinAmountClusterIII = group
                                .Where(x => x.IdNavigation.User.Cluster.Equals(3))
                                .Count(x => (bool)x.Win),
                            CurrencyClusterIII = group
                                .Where(x => x.IdNavigation.User.Cluster.Equals(3))
                                .Sum(x => (bool)x.Win ? x.Currency : 0),
                            USDClusterIII = group
                                .Where(x => x.IdNavigation.User.Cluster.Equals(3))
                                .Sum(x => (bool)x.Win ? x.Currency : 0) * Utilities.GetEventUSDRate(context),
                        }),
                    stageStart => stageStart.Stage,
                    stageEnd => stageEnd.Stage,
                    (stageStart, stageEnd) => new { stageStart, stageEnd })
                .OrderBy(x => x.stageStart.Stage)
                .ToList();

            for (int i = 0; i < stages.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 3)].Value = stages[i].stageStart.Stage;
                worksheet.Cells[String.Concat("B", i + 3)].Value = stages[i].stageStart.StartsClusterO;
                worksheet.Cells[String.Concat("C", i + 3)].Value = stages[i].stageStart.StartsClusterI;
                worksheet.Cells[String.Concat("D", i + 3)].Value = stages[i].stageStart.StartsClusterII;
                worksheet.Cells[String.Concat("E", i + 3)].Value = stages[i].stageStart.StartsClusterIII;
                worksheet.Cells[String.Concat("F", i + 3)].Value = stages[i].stageEnd.EndsClusterO;
                worksheet.Cells[String.Concat("G", i + 3)].Value = stages[i].stageEnd.EndsClusterI;
                worksheet.Cells[String.Concat("H", i + 3)].Value = stages[i].stageEnd.EndsClusterII;
                worksheet.Cells[String.Concat("I", i + 3)].Value = stages[i].stageEnd.EndsClusterIII;
                worksheet.Cells[String.Concat("J", i + 3)].Value = stages[i].stageEnd.WinAmountClusterO;
                worksheet.Cells[String.Concat("K", i + 3)].Value = stages[i].stageEnd.WinAmountClusterI;
                worksheet.Cells[String.Concat("L", i + 3)].Value = stages[i].stageEnd.WinAmountClusterII;
                worksheet.Cells[String.Concat("M", i + 3)].Value = stages[i].stageEnd.WinAmountClusterIII;
                worksheet.Cells[String.Concat("N", i + 3)].Value = stages[i].stageEnd.CurrencyClusterO;
                worksheet.Cells[String.Concat("O", i + 3)].Value = stages[i].stageEnd.CurrencyClusterI;
                worksheet.Cells[String.Concat("P", i + 3)].Value = stages[i].stageEnd.CurrencyClusterII;
                worksheet.Cells[String.Concat("Q", i + 3)].Value = stages[i].stageEnd.CurrencyClusterIII;
                worksheet.Cells[String.Concat("R", i + 3)].Value = stages[i].stageEnd.USDClusterO;
                worksheet.Cells[String.Concat("S", i + 3)].Value = stages[i].stageEnd.USDClusterI;
                worksheet.Cells[String.Concat("T", i + 3)].Value = stages[i].stageEnd.USDClusterII;
                worksheet.Cells[String.Concat("U", i + 3)].Value = stages[i].stageEnd.USDClusterIII;
            }

            Console.WriteLine("Step-by-step by cluster statistics added");

            return excelPackage;
        }
    }
}