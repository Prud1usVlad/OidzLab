using DataAcquisition.Models;
using DataAcquisition.Util;
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
        public static ExcelPackage AddCurrencyRateStatisticsSheet(this ExcelPackage excelPackage, PostgresContext context)
        {

            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Currency rate");

            worksheet.Cells["A1"].Value = "Day";
            worksheet.Cells["B1"].Value = "Rate, $ / curr";

            var data = context.Events
                .Where(e => e.Type == 6)
                .GroupBy(e => e.Date)
                .Select(group => group.Key)
                .ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value = data[i].Value;
                worksheet.Cells[String.Concat("B", i + 2)].Value = 
                    Utilities.GetCurrencyRate(context, data[i]);
            }

            return excelPackage;
        }
    }
}
