using DataAcquisition.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DataAcquisition.Features
{
    public static partial class UserMetrics
    {
        public static ExcelPackage AddDauStatisticsSheet(this ExcelPackage excelPackage, PostgresContext context)
        {

            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("DAU statistics");

            worksheet.Cells["A1"].Value = "Date";
            worksheet.Cells["B1"].Value = "DAU";

            var data = context.Events
                .GroupBy(e => e.Date)
                .Select(group => new
                {
                    Date = group.Key.Value,
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
