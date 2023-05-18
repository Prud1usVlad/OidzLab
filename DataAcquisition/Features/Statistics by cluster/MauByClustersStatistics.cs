using DataAcquisition.Models;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_clusters
{
    public static class MauByClusersStatistics
    {
        public static ExcelPackage AddMauByClustersStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("MAU by clusers statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("MAU by clusers statistics");

            worksheet.Cells["A1"].Value = "Month";
            worksheet.Cells["B1"].Value = "MAU cluster O";
            worksheet.Cells["C1"].Value = "MAU cluster I";
            worksheet.Cells["D1"].Value = "MAU cluster II";
            worksheet.Cells["E1"].Value = "MAU cluster III";

            var data = context.Events
                .GroupBy(e => new DateOnly(e.Date.Value.Year, e.Date.Value.Month, 1))
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
                .ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value = data[i].Date.ToString();
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].ClusterO;
                worksheet.Cells[String.Concat("C", i + 2)].Value = data[i].ClusterI;
                worksheet.Cells[String.Concat("D", i + 2)].Value = data[i].ClusterII;
                worksheet.Cells[String.Concat("E", i + 2)].Value = data[i].ClusterIII;
            }
            
            Console.WriteLine("Mau by clusers statistics added");
            
            return excelPackage;
        }
    }
}