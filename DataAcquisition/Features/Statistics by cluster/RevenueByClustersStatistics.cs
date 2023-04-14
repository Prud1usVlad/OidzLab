using DataAcquisition.Models;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_clusters
{
    public static class RevenueByClustersStatistics
    {
        public static ExcelPackage AddRevenueByClustersStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Revenue by clusters statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Revenue by clusters statistics");

            worksheet.Cells["A1"].Value = "Day";
            worksheet.Cells["B1"].Value = "Revenue cluster O, $";
            worksheet.Cells["C1"].Value = "Revenue cluster I, $";
            worksheet.Cells["D1"].Value = "Revenue cluster II, $";
            worksheet.Cells["E1"].Value = "Revenue cluster III, $";

            var data = context.Events
                .Where(e => e.Type == 6)
                .GroupBy(e => e.Date)
                .Select(group => new
                {
                    Date = group.Key,
                    RevenueClusterO = group
                        .Where(x=> x.User.Cluster.Equals(0))
                        .Sum(i => i.CurrencyPurchase.Price),

                    RevenueClusterI = group
                        .Where(x => x.User.Cluster.Equals(1))
                        .Sum(i => i.CurrencyPurchase.Price),

                    RevenueClusterII = group
                        .Where(x => x.User.Cluster.Equals(2))
                        .Sum(i => i.CurrencyPurchase.Price),

                    RevenueClusterIII = group
                        .Where(x => x.User.Cluster.Equals(3))
                        .Sum(i => i.CurrencyPurchase.Price),
                })
                .OrderBy(x=>x.Date)
                .ToList();

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value =
                    DateOnly.FromDateTime(data[i].Date.Value).ToString();
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].RevenueClusterO;
                worksheet.Cells[String.Concat("C", i + 2)].Value = data[i].RevenueClusterI;
                worksheet.Cells[String.Concat("D", i + 2)].Value = data[i].RevenueClusterII;
                worksheet.Cells[String.Concat("E", i + 2)].Value = data[i].RevenueClusterIII;
            }
            
            Console.WriteLine("Revenue by clusters statistics added");
            
            return excelPackage;
        }
    }
}