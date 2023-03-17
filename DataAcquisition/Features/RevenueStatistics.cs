using DataAcquisition.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAcquisition.Features
{
    public static partial class CurrencyMetrics
    {
        public static ExcelPackage AddRevenueStatisticsSheet(this ExcelPackage excelPackage, PostgresContext context)
        {

            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Revenue statistics");

            worksheet.Cells["A1"].Value = "Day";
            worksheet.Cells["B1"].Value = "Revenue, $";

            var data = context.Events
                .Where(e => e.Type == 6)
                .GroupBy(e => e.Date)
                .Select(group => new
                {
                    Date = group.Key,
                    Revenue = group.Sum(i => i.CurrencyPurchase.Price),
                }).ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value = data[i].Date;
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].Revenue;
            }

            return excelPackage;
        }
    }
}
