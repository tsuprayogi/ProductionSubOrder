using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;
using ProductionOrderAddOn.Models;

namespace ProductionOrderAddOn.Services
{
    internal static class ExcelImportService
    {
        /// <summary>
        /// Import Production Order dari file Excel.
        /// Tiap baris (mulai baris 2) dibuat Production Order baru.
        /// </summary>
        /// <param name="filePath">Path penuh file .xlsx</param>
        /// <returns>Jumlah dokumen berhasil.</returns>
        public static List<ProductionOrderModel> ImportProductionOrders(
            string filePath,
            DateTime? fromDate,
            DateTime? toDate)
        {
            var results = new List<ProductionOrderModel>();

            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    IXLWorksheet ws = workbook.Worksheet(1);

                    if (fromDate.HasValue)
                    {
                        // C3 = baris‑3, kolom‑3  ➜ XLAddress(3,3)
                        string c3Text = ws.Cell(3, 3).GetValue<string>()?.Trim();

                        if (!DateTime.TryParse(c3Text, out DateTime c3Date))
                            throw new Exception($"Header C3 ('{c3Text}') bukan tanggal yang valid.");

                        if (c3Date.Month != fromDate.Value.Month || c3Date.Year != fromDate.Value.Year)
                        {
                            throw new Exception(
                            $"The month and year in Excel header (C3: {c3Date:MM/yyyy}) " +
                            $"do not match the selected Date From ({fromDate:MM/yyyy}).");
                        }
                    }

                    if (toDate.HasValue)
                    {
                        // C3 = baris‑3, kolom‑3  ➜ XLAddress(3,3)
                        string c3Text = ws.Cell(3, 3).GetValue<string>()?.Trim();

                        if (!DateTime.TryParse(c3Text, out DateTime c3Date))
                            throw new Exception($"Header C3 ('{c3Text}') bukan tanggal yang valid.");

                        if (c3Date.Month != toDate.Value.Month || c3Date.Year != toDate.Value.Year)
                        {
                            throw new Exception(
                            $"The month and year in Excel header (C3: {c3Date:MM/yyyy}) " +
                            $"do not match the selected Date To ({toDate:MM/yyyy}).");
                        }
                    }
                    
                    const int startColDate = 3;   // kolom C
                    const int endCol = 34;  // kolom AH
                    
                    // 1️⃣  Tentukan rentang efektif
                    DateTime rangeStart = (fromDate ?? DateTime.MinValue).Date;
                    DateTime rangeEnd = (toDate ?? DateTime.MaxValue).Date;
                    
                    // 2️⃣  Petakan kolom‑>tanggal di header (baris 3)
                    var colDateMap = new List<(int ColIndex, DateTime OrderDate)>();

                    for (int col = startColDate; col <= endCol; col++)
                    {
                        string headerText = ws.Cell(3, col).GetValue<string>();

                        if (DateTime.TryParse(headerText, out DateTime headerDate))
                        {
                            headerDate = headerDate.Date; // abaikan komponen waktu
                            if (headerDate >= rangeStart && headerDate <= rangeEnd)
                            {
                                colDateMap.Add((col, headerDate));
                            }
                        }
                    }

                    bool isSameMonth = colDateMap
                    .Select(x => new { x.OrderDate.Year, x.OrderDate.Month })
                    .Distinct()
                    .Count() == 1;

                    if (!isSameMonth)
                    {
                        throw new Exception("Dates must be in 1 month period per file");
                    }


                    // 3️⃣  Loop baris data (mulai baris ke‑4)
                    foreach (IXLRow row in ws.RowsUsed().Skip(3))
                    {
                        foreach (var (colIndex, orderDate) in colDateMap)
                        {
                            IXLCell cell = row.Cell(colIndex);

                            if (!cell.IsEmpty() && !string.IsNullOrWhiteSpace(cell.GetString()) &&
                                double.TryParse(cell.GetValue<string>(), out double qty))
                            {
                                results.Add(new ProductionOrderModel
                                {
                                    ProdNo = row.Cell(1).GetValue<string>().Trim(),
                                    ProdDesc = row.Cell(2).GetValue<string>().Trim(),
                                    Qty = qty,
                                    OrderDate = orderDate,
                                    ProdType = ProductionType.FG,
                                });
                            }
                        }
                    }
                    if (results.Any())
                    {
                        // Find duplicates
                        var duplicateIds = results
                            .GroupBy(i => new { i.ProdNo, i.OrderDate })
                            .Where(g => g.Count() > 1)
                            .Select(g => $"{g.Key.ProdNo} (Order Date: {g.Key.OrderDate:yyyy-MM-dd})")
                            .ToList();

                        if (duplicateIds.Any())
                        {
                            throw new Exception ("Duplicates found for Item: " + string.Join(", ", duplicateIds));
                        }

                        results = results.OrderBy(r => r.OrderDate).ToList();
                    }
                }
            }
            catch (Exception ex)
            {
                // Bungkus ulang supaya caller dapat pesan kontekstual
                throw new Exception($"Failed to proceed Excel file: {ex.Message}", ex);
            }

            return results;
        }

    }

}
