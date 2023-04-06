using DataAcquisition.Models;
using OfficeOpenXml;

namespace DataAcquisition.Features.Statistics_by_genders
{
    public static class NewUsersByGenderStatistics
    {
        public static ExcelPackage AddNewUsersByGenderStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("New users by gender statistics initialized");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("New Users by gender");

            worksheet.Cells["A1"].Value = "Day";
            worksheet.Cells["B1"].Value = "Male users";
            worksheet.Cells["C1"].Value = "Female users";
            worksheet.Cells["D1"].Value = "Male users total amount";
            worksheet.Cells["E1"].Value = "Female users total amount";

            var data = context.Events
                .Where(i => i.Type == 2)
                .GroupBy(e => e.Date)
                .Select(group => new
                {
                    Date = group.Key,
                    MaleUsers = group.Count(x => x.User.Gender.Equals("male")),
                    FemaleUsers = group.Count(x => x.User.Gender.Equals("female"))
                })
                .OrderBy(x=>x.Date)
                .ToList();

            
            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value =
                    DateOnly.FromDateTime(data[i].Date.Value).ToString();
                worksheet.Cells[String.Concat("B", i + 2)].Value = data[i].MaleUsers;
                worksheet.Cells[String.Concat("C", i + 2)].Value = data[i].FemaleUsers;
            }
            
            worksheet.Cells[String.Concat("D", 2)].Value = data.Sum(x=>x.MaleUsers);
            worksheet.Cells[String.Concat("E", 2)].Value = data.Sum(x=>x.FemaleUsers);
            
            Console.WriteLine("New users by gender statistics added");
            
            return excelPackage;
        }
    }
}