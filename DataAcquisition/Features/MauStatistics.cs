using DataAcquisition.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAcquisition.Features
{
    public static partial class UserMetrics
    {
        public static ExcelPackage AddMauStatistics(this ExcelPackage excelPackage, PostgresContext context)
        {

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
                worksheet.Cells[String.Concat("A", i + 2)].Value = data[i].Date;
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].Users;
            }

            return excelPackage;
        }
    }
}
