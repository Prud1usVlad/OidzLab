using DataAcquisition.Models;
using DataAcquisition.Util;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_age
{
    public static class DauByAgeStatistics
    {
        public static ExcelPackage AddDauByAgeStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("DAU by age statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("DAU by age statistics");

            var groupsAmount = 6;
            var ages = Utilities.GetAgeGroups(context, groupsAmount - 1);

            worksheet.Cells["A1"].Value = "Date";
            for (int i = 0; i < groupsAmount; i++)
            {
                worksheet.Cells[String.Concat(Utilities.GetCellColumnAddress(i + 2), "1")]
                    .Value = i == 0
                    ? String.Concat(0.ToString(), " - ", (ages[i] - 1).ToString())
                    : i + 1 < groupsAmount
                        ? String.Concat(ages[i - 1].ToString(), " - ", (ages[i] - 1).ToString())
                        : String.Concat(ages[i - 1].ToString(), "+");
            }

            var data = context.Events
                .GroupBy(e => e.Date)
                .Select(group => new
                {
                    Date = group.Key,
                    Users1 = 0,
                    Users2 = group
                        .GroupBy(o => o.UserId)
                        .Count(x => x.Any(y => 
                            ages[0] <= y.User.Age && 
                            y.User.Age < ages[1])),
                    Users3 = group
                        .GroupBy(o => o.UserId)
                        .Count(x => x.Any(y => 
                            ages[1] <= y.User.Age && 
                            y.User.Age < ages[2])),
                    Users4 = group
                        .GroupBy(o => o.UserId)
                        .Count(x => x.Any(y => 
                            ages[2] <= y.User.Age && 
                            y.User.Age < ages[3])),
                    Users5 = group
                        .GroupBy(o => o.UserId)
                        .Count(x => x.Any(y => 
                            ages[3] <= y.User.Age && 
                            y.User.Age < ages[4])),
                    Users6 = 0
                })
                .OrderBy(x=> x.Date.ToString())
                .ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value =
                    DateOnly.FromDateTime(data[i].Date.Value).ToString();
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].Users1;
                worksheet.Cells[String.Concat("C", i + 2)].Value = data[i].Users2;
                worksheet.Cells[String.Concat("D", i + 2)].Value = data[i].Users3;
                worksheet.Cells[String.Concat("E", i + 2)].Value = data[i].Users4;
                worksheet.Cells[String.Concat("F", i + 2)].Value = data[i].Users5;
                worksheet.Cells[String.Concat("G", i + 2)].Value = data[i].Users6;
            }

            Console.WriteLine("DAU by age statistics added");
            
            return excelPackage;
        }
    }
}