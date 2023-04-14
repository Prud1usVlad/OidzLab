using DataAcquisition.Models;
using DataAcquisition.Util;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_clusters
{
    public static class PreliminaryByClustersStatistics
    {
        public static ExcelPackage AddPreliminaryByClustersStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Preliminary by clusters statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Preliminary by clusters statistics");

            worksheet.Cells["A1"].Value = "Item name";
            worksheet.Cells["A1:A2"].Merge = true;

            worksheet.Cells["B1"].Value = "Item amount";
            worksheet.Cells["B1:E1"].Merge = true;
            worksheet.Cells["B1"].Value = "Cluster O";
            worksheet.Cells["C1"].Value = "Cluster I";
            worksheet.Cells["D1"].Value = "Cluster II";
            worksheet.Cells["E1"].Value = "Cluster III";

            worksheet.Cells["F1"].Value = "Currency";
            worksheet.Cells["F1:I1"].Merge = true;
            worksheet.Cells["F1"].Value = "Cluster O";
            worksheet.Cells["G1"].Value = "Cluster I";
            worksheet.Cells["H1"].Value = "Cluster II";
            worksheet.Cells["I1"].Value = "Cluster III";

            worksheet.Cells["J1"].Value = "USD";
            worksheet.Cells["J1:M1"].Merge = true;
            worksheet.Cells["J1"].Value = "Cluster O";
            worksheet.Cells["K1"].Value = "Cluster I";
            worksheet.Cells["L1"].Value = "Cluster II";
            worksheet.Cells["M1"].Value = "Cluster III";

            var items = context.ItemPurchases
                .GroupBy(purchase => purchase.ItemName)
                .Select(group =>
                    new
                    {
                        ItemName = group.Key,
                        ItemAmountClusterO = group
                            .Count(x => x.IdNavigation.User.Cluster.Equals(0)),
                        CurrencyClusterO = group
                            .Where(x => x.IdNavigation.User.Cluster.Equals(0))
                            .Sum(x => x.Price),
                        USDClusterO = group
                            .Where(x => x.IdNavigation.User.Cluster.Equals(0))
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),

                        ItemAmountClusterI = group
                            .Count(x => x.IdNavigation.User.Cluster.Equals(1)),
                        CurrencyClusterI = group
                            .Where(x => x.IdNavigation.User.Cluster.Equals(1))
                            .Sum(x => x.Price),
                        USDClusterI = group
                            .Where(x => x.IdNavigation.User.Cluster.Equals(1))
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),

                        ItemAmountClusterII = group
                            .Count(x => x.IdNavigation.User.Cluster.Equals(2)),
                        CurrencyClusterII = group
                            .Where(x => x.IdNavigation.User.Cluster.Equals(2))
                            .Sum(x => x.Price),
                        USDClusterII = group
                            .Where(x => x.IdNavigation.User.Cluster.Equals(2))
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),

                        ItemAmountClusterIII = group
                            .Count(x => x.IdNavigation.User.Cluster.Equals(3)),
                        CurrencyClusterIII = group
                            .Where(x => x.IdNavigation.User.Cluster.Equals(3))
                            .Sum(x => x.Price),
                        USDClusterIII = group
                            .Where(x => x.IdNavigation.User.Cluster.Equals(3))
                            .Sum(x => x.Price) * Utilities.GetEventUSDRate(context),
                    })
                .OrderBy(x=>x.ItemName)
                .ToList();

            for (int i = 0; i < items.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 3)].Value = items[i].ItemName;
                worksheet.Cells[String.Concat("B", i + 3)].Value = items[i].ItemAmountClusterO;
                worksheet.Cells[String.Concat("C", i + 3)].Value = items[i].ItemAmountClusterI;
                worksheet.Cells[String.Concat("D", i + 3)].Value = items[i].ItemAmountClusterII;
                worksheet.Cells[String.Concat("E", i + 3)].Value = items[i].ItemAmountClusterIII;
                worksheet.Cells[String.Concat("F", i + 3)].Value = items[i].CurrencyClusterO;
                worksheet.Cells[String.Concat("G", i + 3)].Value = items[i].CurrencyClusterI;
                worksheet.Cells[String.Concat("H", i + 3)].Value = items[i].CurrencyClusterII;
                worksheet.Cells[String.Concat("I", i + 3)].Value = items[i].CurrencyClusterIII;
                worksheet.Cells[String.Concat("J", i + 3)].Value = items[i].USDClusterO;
                worksheet.Cells[String.Concat("K", i + 3)].Value = items[i].USDClusterI;
                worksheet.Cells[String.Concat("L", i + 3)].Value = items[i].USDClusterII;
                worksheet.Cells[String.Concat("M", i + 3)].Value = items[i].USDClusterIII;
            }

            Console.WriteLine("Preliminary by clusters statistics added");
            
            return excelPackage;
        }
    }
}