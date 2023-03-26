using DataAcquisition.Models;
using OfficeOpenXml;

namespace DataAcquisition.Features
{
    public static partial class UserMetrics
    {
        public static ExcelPackage AddMauStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("MAU statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("MAU statistics");

            worksheet.Cells["A1"].Value = "Month";
            worksheet.Cells["B1"].Value = "MAU";

            var data = context.Events
                .GroupBy(e => new DateOnly(e.Date.Value.Year, e.Date.Value.Month, 1))
                .Select(group => new
                {
                    Date = group.Key,
                    Users = group.GroupBy(o => o.UserId).Count()
                }).ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value = data[i].Date.ToString();
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].Users;
            }
            Console.WriteLine("Mau statistics added");
            return excelPackage;
        }
    }
}
