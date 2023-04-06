using DataAcquisition.Models;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_genders
{
    public static class MauByGenderStatistics
    {
        public static ExcelPackage AddMauByGenderStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("MAU by gender statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("MAU by gender statistics");

            worksheet.Cells["A1"].Value = "Month";
            worksheet.Cells["B1"].Value = "MAU male";
            worksheet.Cells["C1"].Value = "MAU female";

            var data = context.Events
                .GroupBy(e => new DateOnly(e.Date.Value.Year, e.Date.Value.Month, 1))
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
                .ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value = data[i].Date.ToString();
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].MaleUsers;
                worksheet.Cells[String.Concat("C", i + 2)].Value = data[i].FemaleUsers;
            }
            
            Console.WriteLine("Mau by gender statistics added");
            
            return excelPackage;
        }
    }
}