using DataAcquisition.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAcquisition.Features.Statistics_by_cheaters
{
    public static class DauByCheatersStatistics
    {

        public static ExcelPackage AddDauByCheatersStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("DAU by cheaters statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("DAU by cheaters statistics");

            worksheet.Cells["A1"].Value = "Date";
            worksheet.Cells["B1"].Value = "DAU cheaters";
            worksheet.Cells["C1"].Value = "DAU non cheaters";
            var data = context.Events
                .GroupBy(e => e.Date)
                .Select(group => new
                {
                    Date = group.Key,
                    Cheaters = group
                        .GroupBy(o => o.UserId)
                        .Count(x => x.Any(y => y.User.IsCheater.Equals(true))),
                    NonCheaters = group
                        .GroupBy(o => o.UserId)
                        .Count(x => x.Any(y => y.User.IsCheater.Equals(false)))
                })
                .OrderBy(x => x.Date.ToString())
                .ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value =
                    DateOnly.FromDateTime(data[i].Date.Value).ToString();
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].Cheaters;
                worksheet.Cells[String.Concat("C", i + 2)].Value = data[i].NonCheaters;
            }

            Console.WriteLine("DAU by cheaters statistics added");

            return excelPackage;
        }

    }
}
