using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_age
{
    public static class NewUsersByAgeStatistics
    {
        public static ExcelPackage AddNewUsersByAgeStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("New users by age statistics initialized");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("New Users by age");

            
            var groupsAmount = 6;
            var ages = Utilities.GetAgeGroups(context, groupsAmount - 1);
            
            worksheet.Cells["A1"].Value = "Day";
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
                .Where(i => i.Type == 2)
                .GroupBy(e => e.Date)
                .Select(group => new
                {
                    Date = group.Key,
                    Users1 = 0,
                    Users2 = group
                        .Count(x => ages[0] <= x.User.Age && x.User.Age < ages[1]),
                    Users3 = group
                        .Count(x => ages[1] <= x.User.Age && x.User.Age < ages[2]),
                    Users4 = group
                        .Count(x => ages[2] <= x.User.Age && x.User.Age < ages[3]),
                    Users5 = group
                        .Count(x => ages[3] <= x.User.Age && x.User.Age < ages[4]),
                    Users6 = 0
                })
                .OrderBy(x=>x.Date)
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

            Console.WriteLine("New users by age statistics added");
            
            return excelPackage;
        }
    }
}