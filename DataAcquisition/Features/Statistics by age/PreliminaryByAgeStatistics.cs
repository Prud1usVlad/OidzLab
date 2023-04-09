using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_age
{
    public static class PreliminaryByAgeStatistics
    {
        public static ExcelPackage AddPreliminaryByAgeStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Preliminary by age statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Preliminary by age statistics");

            var groupsAmount = 6;
            var ages = Utilities.GetAgeGroups(context, groupsAmount - 1);
            
            worksheet.Cells["A1"].Value = "Item name";
            worksheet.Cells["A1:A2"].Merge = true;
            worksheet.Cells["B1"].Value = "Item amount";
            worksheet.Cells["B1:C1"].Merge = true;
            for (int i = 0; i < groupsAmount; i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(i + 2), "2")]
                    .Value = i == 0
                    ? String.Concat(0.ToString(), " - ", (ages[i] - 1).ToString())
                    : i + 1 < groupsAmount
                        ? String.Concat(ages[i - 1].ToString(), " - ", (ages[i] - 1).ToString())
                        : String.Concat(ages[i - 1].ToString(), "+");
            }
            worksheet.Cells["H1"].Value = "Currency";
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
            worksheet.Cells["N1"].Value = "USD";
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

            var items = context.ItemPurchases
                .GroupBy(purchase => purchase.ItemName)
                .Select(group =>
                    new
                    {
                        ItemName = group.Key,
                        ItemAmount1 = 0,
                        ItemAmount2 = group
                            .Count(x => ages[0] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[1]),
                        ItemAmount3 = group
                            .Count(x => ages[1] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[2]),
                        ItemAmount4 = group
                            .Count(x => ages[2] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[3]),
                        ItemAmount5 = group
                            .Count(x => ages[3] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[4]),
                        ItemAmount6= 0,
                        Currency1 = 0,
                        Currency2 = group
                            .Where(x => ages[0] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[1])
                            .Sum(x => x.Price),
                        Currency3 = group
                            .Where(x => ages[1] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[2])
                            .Sum(x => x.Price),
                        Currency4 = group
                            .Where(x => ages[2] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[3])
                            .Sum(x => x.Price),
                        Currency5 = group
                            .Where(x => ages[3] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[4])
                            .Sum(x => x.Price),
                        Currency6 = 0,
                        USD1 = 0,
                        USD2 = group
                            .Where(x => ages[0] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[1])
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),
                        USD3 = group
                            .Where(x => ages[1] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[2])
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),
                        USD4 = group
                            .Where(x => ages[2] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[3])
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),
                        USD5 = group
                            .Where(x => ages[3] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[4])
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),
                        USD6 = 0
                    })
                .OrderBy(x=>x.ItemName)
                .ToList();

            for (int i = 0; i < items.Count(); i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(1), (i + 3).ToString())]
                    .Value = items[i].ItemName;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(2), (i + 3).ToString())]
                    .Value = items[i].ItemAmount1;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(3), (i + 3).ToString())]
                    .Value = items[i].ItemAmount2;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(4), (i + 3).ToString())]
                    .Value = items[i].ItemAmount3;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(5), (i + 3).ToString())]
                    .Value = items[i].ItemAmount4;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(6), (i + 3).ToString())]
                    .Value = items[i].ItemAmount5;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(7), (i + 3).ToString())]
                    .Value = items[i].ItemAmount6;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(8), (i + 3).ToString())]
                    .Value = items[i].Currency1;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(9), (i + 3).ToString())]
                    .Value = items[i].Currency2;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(10), (i + 3).ToString())]
                    .Value = items[i].Currency3;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(11), (i + 3).ToString())]
                    .Value = items[i].Currency4;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(12), (i + 3).ToString())]
                    .Value = items[i].Currency5;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(13), (i + 3).ToString())]
                    .Value = items[i].Currency6;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(14), (i + 3).ToString())]
                    .Value = items[i].USD1;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(15), (i + 3).ToString())]
                    .Value = items[i].USD2;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(16), (i + 3).ToString())]
                    .Value = items[i].USD3;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(17), (i + 3).ToString())]
                    .Value = items[i].USD4;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(18), (i + 3).ToString())]
                    .Value = items[i].USD5;
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(19), (i + 3).ToString())]
                    .Value = items[i].USD6;
            }

            Console.WriteLine("Preliminary by age statistics added");
            
            return excelPackage;
        }
    }
}