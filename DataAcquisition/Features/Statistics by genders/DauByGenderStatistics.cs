using DataAcquisition.Models;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_genders
{
    public static class DauByGenderStatistics
    {
        public static ExcelPackage AddDauByGenderStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("DAU by gender statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("DAU by gender statistics");

            worksheet.Cells["A1"].Value = "Date";
            worksheet.Cells["B1"].Value = "DAU male";
            worksheet.Cells["C1"].Value = "DAU female";
            var data = context.Events
                .GroupBy(e => e.Date)
                .Select(group => new
                {
                    Date = group.Key,
                    MaleUsers = group
                        .GroupBy(o => o.UserId)
                        .Count(x => x.Any(y => y.User.Gender.Equals("male"))),
                    FemaleUsers = group
                        .GroupBy(o => o.UserId)
                        .Count(x => x.Any(y => y.User.Gender.Equals("female")))
                })
                .OrderBy(x=> x.Date.ToString())
                .ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value =
                    DateOnly.FromDateTime(data[i].Date.Value).ToString();
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].MaleUsers;
                worksheet.Cells[String.Concat("C", i + 2)].Value = data[i].FemaleUsers;
            }

            Console.WriteLine("DAU by gender statistics added");
            return excelPackage;
        }
    }
}