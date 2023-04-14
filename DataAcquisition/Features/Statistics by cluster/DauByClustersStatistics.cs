using DataAcquisition.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAcquisition.Features.Statistics_by_clusters
{
    public static class DauByClustersStatistics
    {

        public static ExcelPackage AddDauByClustersStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("DAU by clusters statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("DAU by clusters statistics");

            worksheet.Cells["A1"].Value = "Date";
            worksheet.Cells["B1"].Value = "DAU cluster O";
            worksheet.Cells["C1"].Value = "DAU cluster I";
            worksheet.Cells["D1"].Value = "DAU cluster II";
            worksheet.Cells["E1"].Value = "DAU cluster III";
            var data = context.Events
                .GroupBy(e => e.Date)
                .Select(group => new
                {
                    Date = group.Key,
                    ClusterO = group
                        .GroupBy(o => o.UserId)
                        .Count(x => x.Any(y => y.User.Cluster.Equals(0))),
                    ClusterI = group
                        .GroupBy(o => o.UserId)
                        .Count(x => x.Any(y => y.User.Cluster.Equals(1))),
                    ClusterII = group
                        .GroupBy(o => o.UserId)
                        .Count(x => x.Any(y => y.User.Cluster.Equals(2))),
                    ClusterIII = group
                        .GroupBy(o => o.UserId)
                        .Count(x => x.Any(y => y.User.Cluster.Equals(3)))
                })
                .OrderBy(x => x.Date.ToString())
                .ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value =
                    DateOnly.FromDateTime(data[i].Date.Value).ToString();
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].ClusterO;
                worksheet.Cells[String.Concat("C", i + 2)].Value = data[i].ClusterI;
                worksheet.Cells[String.Concat("D", i + 2)].Value = data[i].ClusterII;
                worksheet.Cells[String.Concat("E", i + 2)].Value = data[i].ClusterIII;
            }

            Console.WriteLine("DAU by clusters statistics added");

            return excelPackage;
        }

    }
}
