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
        public static ExcelPackage AddNewUsersStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {

            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("New Users");

            worksheet.Cells["A1"].Value = "Day";
            worksheet.Cells["B1"].Value = "Users";

            var data = context.Events
                .GroupBy(e => e.Date)
                .Select(group => new
                {
                    Date = group.Key,
                    Users = group.Where(i => i.Type == 2).Count(),
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
