﻿using OfficeOpenXml;
using DataAcquisition.Models;
using DataAcquisition.Util;

namespace DataAcquisition.Features
{
    public static class PreliminaryStatistics
    {
        public static ExcelPackage AddPreliminaryStatisticsSheet(this ExcelPackage excelPackage, OidzDbContext context)
        {
            Console.WriteLine("Preliminary statistics init");
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Preliminary statistics");

            worksheet.Cells["A1"].Value = "Item name";
            worksheet.Cells["B1"].Value = "Item amount";
            worksheet.Cells["C1"].Value = "Currency";
            worksheet.Cells["D1"].Value = "USD";
            
            var items = context.ItemPurchases
                .GroupBy(purchase => purchase.ItemName)
                .Select(group =>
                    new
                    {
                        ItemName = group.Key,
                        ItemAmount = group.Count(),
                        Currency = group.Sum(x => x.Price),
                        USD = group.Sum(x => x.Price) * Utilities.GetEventUSDRate(context)
                    })
                .OrderBy(x=>x.ItemName)
                .ToList();

            for (int i = 0; i < items.Count(); i++)
            {
                worksheet.Cells[String.Concat("A", i + 2)].Value = items[i].ItemName;
                worksheet.Cells[String.Concat("B", i + 2)].Value = items[i].ItemAmount;
                worksheet.Cells[String.Concat("C", i + 2)].Value = items[i].Currency;
                worksheet.Cells[String.Concat("D", i + 2)].Value = items[i].USD;
            }

            Console.WriteLine("Preliminary statistics added");
            return excelPackage;
        }
    }
}