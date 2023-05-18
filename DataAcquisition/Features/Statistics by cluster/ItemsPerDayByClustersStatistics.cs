using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_clusters
{
    public static class ItemsPerDayByClustersStatistics
    {
        public static ExcelPackage AddItemsPerDayByClustersStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Items-per-day by clusters statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Items-per-day by clusters statistics");

            worksheet.Cells["A1"].Value = "Date";
            worksheet.Cells["A1:A2"].Merge = true;
            
            worksheet.Cells["B1"].Value = "Items amount";
            worksheet.Cells["B1:E1"].Merge = true;
            worksheet.Cells["B2"].Value = "Cluster O";
            worksheet.Cells["C2"].Value = "Cluster I";
            worksheet.Cells["D2"].Value = "Cluster II";
            worksheet.Cells["E2"].Value = "Cluster III";
            
            worksheet.Cells["F1"].Value = "USD";
            worksheet.Cells["F1:I1"].Merge = true;
            worksheet.Cells["F2"].Value = "Cluster O";
            worksheet.Cells["G2"].Value = "Cluster I";
            worksheet.Cells["H2"].Value = "Cluster II";
            worksheet.Cells["I2"].Value = "Cluster III";

            var items = context.ItemPurchases
                .GroupBy(purchase => purchase.IdNavigation.Date)
                .Select(group =>
                    new
                    {
                        Date = group.Key,
                        ItemAmountClusterO = group
                            .Count(x => x.IdNavigation.User.Cluster.Equals(0)),
                        USDClusterO = group
                            .Where(x => x.IdNavigation.User.Cluster.Equals(0))
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),

                        ItemAmountClusterI = group
                            .Count(x => x.IdNavigation.User.Cluster.Equals(1)),
                        USDClusterI = group
                            .Where(x => x.IdNavigation.User.Cluster.Equals(1))
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),

                        ItemAmountClusterII = group
                            .Count(x => x.IdNavigation.User.Cluster.Equals(2)),
                        USDClusterII = group
                            .Where(x => x.IdNavigation.User.Cluster.Equals(2))
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),

                        ItemAmountClusterIII = group
                            .Count(x => x.IdNavigation.User.Cluster.Equals(3)),
                        USDClusterIII = group
                            .Where(x => x.IdNavigation.User.Cluster.Equals(3))
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),
                    })
                .OrderBy(x=>x.Date)
                .ToList();

            for (int i = 0; i < items.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 3)].Value = DateOnly.FromDateTime(items[i].Date.Value).ToString();
                worksheet.Cells[String.Concat("B", i + 3)].Value = items[i].ItemAmountClusterO;
                worksheet.Cells[String.Concat("C", i + 3)].Value = items[i].ItemAmountClusterI;
                worksheet.Cells[String.Concat("D", i + 3)].Value = items[i].ItemAmountClusterII;
                worksheet.Cells[String.Concat("E", i + 3)].Value = items[i].ItemAmountClusterIII;
                worksheet.Cells[String.Concat("F", i + 3)].Value = items[i].USDClusterO;
                worksheet.Cells[String.Concat("G", i + 3)].Value = items[i].USDClusterI;
                worksheet.Cells[String.Concat("H", i + 3)].Value = items[i].USDClusterII;
                worksheet.Cells[String.Concat("I", i + 3)].Value = items[i].USDClusterIII;
            }

            Console.WriteLine("Items-per-day by clusters statistics added");
            
            return excelPackage;
        }
    }
}