using DataAcquisition.Models;
using DataAcquisition.Util;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_age
{
    public static class ItemsPerDayByAgeStatistics
    {
        public static ExcelPackage AddItemsPerDayByAgeStatisticsSheet(this ExcelPackage excelPackage,
            OidzDbContext context)
        {
            Console.WriteLine("Items-per-day by age statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Items-per-day by age statistics");

            var groupsAmount = 6;
            var ages = Utilities.GetAgeGroups(context, groupsAmount - 1);

            worksheet.Cells["A1"].Value = "Date";
            worksheet.Cells["A1:A2"].Merge = true;
            worksheet.Cells["B1"].Value = "Items amount";
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
            worksheet.Cells["H1"].Value = "USD";
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
            
            var items = context.ItemPurchases
                .GroupBy(purchase => purchase.IdNavigation.Date)
                .Select(group =>
                    new
                    {
                        Date = group.Key,
                        ItemAmount1 = 0,
                        USD1 = 0,
                        ItemAmount2 = group
                            .Count(x => ages[0] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[1]),
                        USD2 = group
                            .Where(x => ages[0] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[1])
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),
                        ItemAmount3 = group
                            .Count(x => ages[1] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[2]),
                        USD3 = group
                            .Where(x => ages[1] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[2])
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),
                        ItemAmount4 = group
                            .Count(x => ages[2] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[3]),
                        USD4 = group
                            .Where(x => ages[2] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[3])
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),
                        ItemAmount5 = group
                            .Count(x => ages[3] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[4]),
                        USD5 = group
                            .Where(x => ages[3] <= x.IdNavigation.User.Age && x.IdNavigation.User.Age < ages[4])
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),
                        ItemAmount6 = 0,
                        USD6 = 0
                    })
                .OrderBy(x => x.Date)
                .ToList();

            for (int i = 0; i < items.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 3)].Value =
                    DateOnly.FromDateTime(items[i].Date.Value).ToString();
                worksheet.Cells[String.Concat("B", i + 3)].Value = items[i].ItemAmount1;
                worksheet.Cells[String.Concat("C", i + 3)].Value = items[i].ItemAmount2;
                worksheet.Cells[String.Concat("D", i + 3)].Value = items[i].ItemAmount3;
                worksheet.Cells[String.Concat("E", i + 3)].Value = items[i].ItemAmount4;
                worksheet.Cells[String.Concat("F", i + 3)].Value = items[i].ItemAmount5;
                worksheet.Cells[String.Concat("G", i + 3)].Value = items[i].ItemAmount6;
                worksheet.Cells[String.Concat("H", i + 3)].Value = items[i].USD1;
                worksheet.Cells[String.Concat("I", i + 3)].Value = items[i].USD2;
                worksheet.Cells[String.Concat("J", i + 3)].Value = items[i].USD3;
                worksheet.Cells[String.Concat("K", i + 3)].Value = items[i].USD4;
                worksheet.Cells[String.Concat("L", i + 3)].Value = items[i].USD5;
                worksheet.Cells[String.Concat("M", i + 3)].Value = items[i].USD6;
            }

            Console.WriteLine("Items-per-day by age statistics added");

            return excelPackage;
        }
    }
}