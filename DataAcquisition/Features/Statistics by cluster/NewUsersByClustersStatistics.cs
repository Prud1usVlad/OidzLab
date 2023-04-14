using DataAcquisition.Models;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_clusters
{
    public static class NewUsersByClusersStatistics
    {
        public static ExcelPackage AddNewUsersByClustersStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("New users by clusters statistics initialized");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("New Users by clusters");

            worksheet.Cells["A1"].Value = "Day";
            worksheet.Cells["B1"].Value = "Cluster O";
            worksheet.Cells["C1"].Value = "Cluster I";
            worksheet.Cells["D1"].Value = "Cluster II";
            worksheet.Cells["E1"].Value = "Cluster III";
            worksheet.Cells["F1"].Value = "Cluster O total";
            worksheet.Cells["G1"].Value = "Cluster I total";
            worksheet.Cells["H1"].Value = "Cluster II total";
            worksheet.Cells["I1"].Value = "Cluster III total";

            var data = context.Events
                .Where(i => i.Type == 2)
                .GroupBy(e => e.Date)
                .Select(group => new
                {
                    Date = group.Key,
                    ClusterO = group.Count(x => x.User.Cluster.Equals(0)),
                    ClusterI = group.Count(x => x.User.Cluster.Equals(1)),
                    ClusterII = group.Count(x => x.User.Cluster.Equals(2)),
                    ClusterIII = group.Count(x => x.User.Cluster.Equals(3)),
                })
                .OrderBy(x=>x.Date)
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
            
            worksheet.Cells[String.Concat("F", 2)].Value = data.Sum(x => x.ClusterO);
            worksheet.Cells[String.Concat("G", 2)].Value = data.Sum(x => x.ClusterI);
            worksheet.Cells[String.Concat("H", 2)].Value = data.Sum(x => x.ClusterII);
            worksheet.Cells[String.Concat("I", 2)].Value = data.Sum(x => x.ClusterIII);

            Console.WriteLine("New users by clusers statistics added");
            
            return excelPackage;
        }
    }
}