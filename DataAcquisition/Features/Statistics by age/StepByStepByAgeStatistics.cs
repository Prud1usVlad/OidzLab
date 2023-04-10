using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_age
{
    public static class StepByStepByAgeStatistics
    {
        public static ExcelPackage AddStepByStepByAgeStatisticsSheet(this ExcelPackage excelPackage,
            OidzDbContext context)
        {
            Console.WriteLine("Step-by-step by age statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Step-by-step by age statistics");

            var groupsAmount = 6;
            var ages = Utilities.GetAgeGroups(context, groupsAmount - 1);
            
            worksheet.Cells["A1"].Value = "Stage";
            worksheet.Cells["A1:A2"].Merge = true;
            worksheet.Cells["B1"].Value = "Starts";
            worksheet.Cells["B1:G1"].Merge = true;
            for (int i = 0; i < groupsAmount; i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(i + 2), "2")]
                    .Value = i == 0
                    ? String.Concat(0.ToString(), " - ", (ages[i] - 1).ToString())
                    : i + 1 < groupsAmount
                        ? String.Concat(ages[i - 1].ToString(), " - ", (ages[i] - 1).ToString())
                        : String.Concat(ages[i - 1].ToString(), "+");
            }
            worksheet.Cells["H1"].Value = "Ends";
            worksheet.Cells["H1:M1"].Merge = true;
            for (int i = 0; i < groupsAmount; i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(i + 8), "2")]
                    .Value = i == 0
                    ? String.Concat(0.ToString(), " - ", (ages[i] - 1).ToString())
                    : i + 1 < groupsAmount
                        ? String.Concat(ages[i - 1].ToString(), " - ", (ages[i] - 1).ToString())
                        : String.Concat(ages[i - 1].ToString(), "+");
            }
            worksheet.Cells["N1"].Value = "Wins";
            worksheet.Cells["N1:S1"].Merge = true;
            for (int i = 0; i < groupsAmount; i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(i + 14), "2")]
                    .Value = i == 0
                    ? String.Concat(0.ToString(), " - ", (ages[i] - 1).ToString())
                    : i + 1 < groupsAmount
                        ? String.Concat(ages[i - 1].ToString(), " - ", (ages[i] - 1).ToString())
                        : String.Concat(ages[i - 1].ToString(), "+");
            }
            worksheet.Cells["T1"].Value = "Currency";
            worksheet.Cells["T1:Y1"].Merge = true;
            for (int i = 0; i < groupsAmount; i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(i + 20), "2")]
                    .Value = i == 0
                    ? String.Concat(0.ToString(), " - ", (ages[i] - 1).ToString())
                    : i + 1 < groupsAmount
                        ? String.Concat(ages[i - 1].ToString(), " - ", (ages[i] - 1).ToString())
                        : String.Concat(ages[i - 1].ToString(), "+");
            }
            worksheet.Cells["Z1"].Value = "USD";
            worksheet.Cells["Z1:AE1"].Merge = true;
            for (int i = 0; i < groupsAmount; i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(i + 26), "2")]
                    .Value = i == 0
                    ? String.Concat(0.ToString(), " - ", (ages[i] - 1).ToString())
                    : i + 1 < groupsAmount
                        ? String.Concat(ages[i - 1].ToString(), " - ", (ages[i] - 1).ToString())
                        : String.Concat(ages[i - 1].ToString(), "+");
            }

            var stages = context.StageStarts
                .GroupBy(stageStart => stageStart.Stage)
                .Select(group => new
                {
                    Stage = group.Key.Value,
                    Starts1 = 0,
                    Starts2 = group.Count(x => ages[0] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[1]),
                    Starts3 = group.Count(x => ages[1] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[2]),
                    Starts4 = group.Count(x => ages[2] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[3]),
                    Starts5 = group.Count(x => ages[3] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[4]),
                    Starts6 = 0
                })
                .Join(
                    context.StageEnds
                        .GroupBy(stageEnd => stageEnd.Stage)
                        .Select(group => new
                        {
                            Stage = group.Key.Value,
                            Ends1 = 0,
                            Ends2 = group
                                .Count(x => ages[0] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[1]),
                            Ends3 = group
                                .Count(x => ages[1] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[2]),
                            Ends4 = group
                                .Count(x => ages[1] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[2]),
                            Ends5 = group
                                .Count(x => ages[3] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[4]),
                            Ends6 = 0,
                            WinAmount1 = 0,
                            WinAmount2 = group
                                .Where(x => ages[0] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[1])
                                .Count(x => (bool)x.Win),
                            WinAmount3 = group
                                .Where(x => ages[1] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[2])
                                .Count(x => (bool)x.Win),
                            WinAmount4 = group
                                .Where(x => ages[2] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[3])
                                .Count(x => (bool)x.Win),
                            WinAmount5 = group
                                .Where(x => ages[3] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[4])
                                .Count(x => (bool)x.Win),
                            WinAmount6 = 0,
                            Currency1 = 0,
                            Currency2 = group
                                .Where(x => ages[0] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[1])
                                .Sum(x => (bool)x.Win ? x.Currency : 0),
                            Currency3 = group
                                .Where(x => ages[1] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[2])
                                .Sum(x => (bool)x.Win ? x.Currency : 0),
                            Currency4 = group
                                .Where(x => ages[2] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[3])
                                .Sum(x => (bool)x.Win ? x.Currency : 0),
                            Currency5 = group
                                .Where(x => ages[3] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[4])
                                .Sum(x => (bool)x.Win ? x.Currency : 0),
                            Currency6 = 0,
                            USD1 = 0,
                            USD2 = group
                                .Where(x => ages[0] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[1])
                                .Sum(x => (bool)x.Win ? x.Currency : 0) * Utilities.GetEventUSDRate(context),
                            USD3 = group
                                .Where(x => ages[1] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[2])
                                .Sum(x => (bool)x.Win ? x.Currency : 0) * Utilities.GetEventUSDRate(context),
                            USD4 = group
                                .Where(x => ages[2] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[3])
                                .Sum(x => (bool)x.Win ? x.Currency : 0) * Utilities.GetEventUSDRate(context),
                            USD5 = group
                                .Where(x => ages[3] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[4])
                                .Sum(x => (bool)x.Win ? x.Currency : 0) * Utilities.GetEventUSDRate(context),
                            USD6 = 0,
                        }),
                    stageStart => stageStart.Stage,
                    stageEnd => stageEnd.Stage,
                    (stageStart, stageEnd) => new { stageStart, stageEnd })
                .OrderBy(x => x.stageStart.Stage)
                .ToList();
            
            for (int i = 0; i < stages.Count(); i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(1), (i + 3).ToString())]
                    .Value = stages[i].stageStart.Stage;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(2), (i + 3).ToString())]
                    .Value = stages[i].stageStart.Starts1;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(3), (i + 3).ToString())]
                    .Value = stages[i].stageStart.Starts2;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(4), (i + 3).ToString())]
                    .Value = stages[i].stageStart.Starts3;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(5), (i + 3).ToString())]
                    .Value = stages[i].stageStart.Starts4;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(6), (i + 3).ToString())]
                    .Value = stages[i].stageStart.Starts5;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(7), (i + 3).ToString())]
                    .Value = stages[i].stageStart.Starts6;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(8), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.Ends1;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(9), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.Ends2;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(10), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.Ends3;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(11), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.Ends4;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(12), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.Ends5;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(13), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.Ends6;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(14), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.WinAmount1;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(15), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.WinAmount2;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(16), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.WinAmount3;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(17), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.WinAmount4;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(18), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.WinAmount5;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(19), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.WinAmount6;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(20), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.Currency1;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(21), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.Currency2;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(22), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.Currency3;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(23), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.Currency4;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(24), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.Currency5;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(25), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.Currency6;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(26), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.USD1;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(27), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.USD2;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(28), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.USD3;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(29), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.USD4;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(30), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.USD5;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(31), (i + 3).ToString())]
                    .Value = stages[i].stageEnd.USD6;
            }

            Console.WriteLine("Step-by-step by age statistics added");

            return excelPackage;
        }
    }
}