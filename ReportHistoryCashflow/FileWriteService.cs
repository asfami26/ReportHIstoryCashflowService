using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml.Wordprocessing;
using ReportHistoryCashflow.Class;
using System.Collections;
using System.Globalization;
using DocumentFormat.OpenXml.Bibliography;

namespace ReportHistoryCashflow
{
    public class FileWriteService : ServiceBase
    {
        public Thread Worker = null;

        public FileWriteService()
        {
            ServiceName = "ReportHIstoryCashflowService";
        }

        protected void Onstart(string[] args)
        {
            ThreadStart start = new ThreadStart(Working);
            Worker = new Thread(start);
            Worker.Start();
        }

        public void Working()
        {
            int nsleep = 1;
            int rowhasil = 0;
            int col = 0;
            int row = 0;
            int rowkeluar = 0;
            try
            {
                while (true)
                {

                    sql conn = new sql();

                    DateTime currentDate = DateTime.Now;
                    DateTime nextYearDate = currentDate.AddYears(1);

                    var workbook = new XLWorkbook();
                    var worksheet = workbook.Worksheets.Add("ReportHistoryCashflow");

                     col = 1;
                     row = 4;

                    var headerCellA2 = worksheet.Cell(2, col);
                    headerCellA2.Value = "Keterangan";
                    headerCellA2.Style.Fill.BackgroundColor = XLColor.TwilightLavender;
                    headerCellA2.Style.Font.FontColor = XLColor.White;
                    col++;

                    while (currentDate <= nextYearDate)
                    {
                        if (currentDate.DayOfWeek != DayOfWeek.Saturday && currentDate.DayOfWeek != DayOfWeek.Sunday)
                        {
                            var headerCell = worksheet.Cell(2, col);
                            headerCell.Value = currentDate.ToString("dd-MMM-yyyy");
                            headerCell.Style.Fill.BackgroundColor = XLColor.TwilightLavender;
                            headerCell.Style.Font.FontColor = XLColor.White;
                            currentDate = currentDate.AddDays(1);
                            col++;
                        }
                        else
                        {
                            currentDate = currentDate.AddDays(1);
                        }
                    }
                    Console.WriteLine("Proses Tanggal Selesai");

                    string sql = "SELECT [Name], [Id], [SortOrder], [ParentId], [SubId] " +
                        "FROM (SELECT Id,[Name],0 AS [Level],ParentKategori_Id AS ParentId, " +
                        "[order],1 AS [SortOrder],NULL AS [SubId] " +
                        "FROM Kategori " +
                        "WHERE Id IN ('3025', '3003', '3006', '3004') " +
                        "UNION ALL " +
                        "SELECT k.Id, CASE WHEN k.Id = '3004' AND k.[Name] = 'Remis' THEN '      ' + k.[Name] " +
                        "ELSE '      ' + k.[Name] END, p.[Level] + 1 AS [Level], k.ParentKategori_Id AS ParentId, k.[order], " +
                        "2 AS [SortOrder], k.Id AS [SubId] FROM Kategori k INNER JOIN (SELECT Id, 0 AS [Level], Id AS ParentId, [order] " +
                        "FROM Kategori " +
                        "WHERE Id IN ('3025', '3003', '3006', '3004')) p ON p.Id = k.ParentKategori_Id) AS T " +
                        "WHERE ParentId IS NULL OR (ParentId <> '3004' OR (ParentId = '3004' AND [id] = '3008')) " +
                        "ORDER BY CASE WHEN ParentId IS NULL THEN [order] ELSE (SELECT [order] FROM Kategori WHERE Id = T.ParentId) END, " +
                        "[SortOrder], [Level], [order]";

                    DataTable dth = conn.GetDataTable(sql);
                    rowhasil = dth.Rows.Count + row;

                    foreach (DataRow dr in dth.Rows)
                    {
                        worksheet.Cell(row, 1).Value = dr["Name"].ToString();
                        string? fbold = dr["SortOrder"].ToString();
                        if (fbold == "1") { worksheet.Cell(row, 1).Style.Font.Bold = true; }
                        else { worksheet.Cell(row, 1).Style.Font.Bold = false; }

                        col = 2;
                        string qtotal = "WITH TanggalProyeksi AS " +
                            "(SELECT DATEADD(DAY, number, CONVERT(date, GETDATE())) AS Tanggal " +
                            "FROM master.dbo.spt_values " +
                            "WHERE type = 'P' " +
                            "AND DATEADD(DAY, number, CONVERT(date, GETDATE())) <= DATEADD(YEAR, 1, CONVERT(date, GETDATE())) " +
                            "AND DATEPART(WEEKDAY, DATEADD(DAY, number, CONVERT(date, GETDATE()))) NOT IN (1, 7)) " +
                            "SELECT CAST(tp.Tanggal AS date) AS Tanggal, pt.Kategori, pt.SubKategori, " +
                            "REPLACE(SUM(pt.Nominal), '.00', '') AS Nominal, " +
                            "(SELECT REPLACE(CAST(SUM(pt2.Nominal) AS varchar), '.00', '') " +
                            "FROM ProTrxFinansial_Log p2 " +
                            "JOIN ProTrxFinansialItem pt2 ON p2.Data_Id = pt2.ProTrxFinansial_Id " +
                            "WHERE p2.TypeTransaksi = '1' " +
                            "AND CAST(p2.TanggalProyeksi AS date) = CAST(tp.Tanggal AS date)) AS TotalKategori " +
                            "FROM TanggalProyeksi tp " +
                            "LEFT JOIN ProTrxFinansial_Log p ON CONVERT(date, p.TanggalProyeksi) = CAST(tp.Tanggal AS date) " +
                            "LEFT JOIN ProTrxFinansialItem pt ON p.Data_Id = pt.ProTrxFinansial_Id " +
                            "WHERE (pt.Kategori IS NULL OR pt.SubKategori IS NULL OR " +
                            "(pt.Kategori = '" + dr["ParentId"].ToString() + "' AND pt.SubKategori = '" + dr["SubId"].ToString() + "')) " +
                            "AND (p.TypeTransaksi = '1' OR p.TypeTransaksi IS NULL) " +
                            "GROUP BY CAST(tp.Tanggal AS date), pt.Kategori, pt.SubKategori " +
                            "ORDER BY CAST(tp.Tanggal AS date)";

                        DataTable dttotal = conn.GetDataTable(qtotal);

                        foreach (DataRow drtotal in dttotal.Rows)
                        {
                            worksheet.Cell(row, col).Value = drtotal["Nominal"].ToString() == null ? "" : drtotal["Nominal"].ToString();
                            if (drtotal["TotalKategori"] != null && !DBNull.Value.Equals(drtotal["TotalKategori"]))
                            {
                                worksheet.Cell(rowhasil, col).Value = drtotal["TotalKategori"].ToString();
                            }

                            col++;
                        }
                        row++;
                    }
                    Console.WriteLine("Proses Pertama Selesai");

                    var headerRange = worksheet.Range(worksheet.Cell(3, 1), worksheet.Cell(3, col - 1));
                    headerRange.Merge();
                    headerRange.Value = "Dana Masuk ";
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Font.FontColor = XLColor.Black;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightPink;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    worksheet.Cell(row, 1).Value = "Total Dana Masuk";
                    headerRange = worksheet.Range(worksheet.Cell(row, 1), worksheet.Cell(row, col - 1));
                    headerRange.Style.Font.FontColor = XLColor.Black;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                    row++;

                    headerRange = worksheet.Range(worksheet.Cell(row, 1), worksheet.Cell(row, col - 1));
                    headerRange.Merge();
                    headerRange.Value = "Dana Keluar";
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Font.FontColor = XLColor.Black;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightPink;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    row++;

                    string sql2 = "SELECT [Name], [Id], [SortOrder], [ParentId], [SubId] " +
                        "FROM (SELECT Id,[Name],0 AS [Level],ParentKategori_Id AS ParentId, " +
                        "[order],1 AS [SortOrder],NULL AS [SubId] " +
                        "FROM Kategori " +
                        "WHERE Id IN ('3025', '3003', '3006', '3004') " +
                        "UNION ALL " +
                        "SELECT k.Id, CASE WHEN k.Id = '3004' AND k.[Name] = 'Supply' THEN '      ' + k.[Name] " +
                        "ELSE '      ' + k.[Name] END, p.[Level] + 1 AS [Level], k.ParentKategori_Id AS ParentId, k.[order], " +
                        "2 AS [SortOrder], k.Id AS [SubId] FROM Kategori k INNER JOIN (SELECT Id, 0 AS [Level], Id AS ParentId, [order] " +
                        "FROM Kategori " +
                        "WHERE Id IN ('3025', '3003', '3006', '3004')) p ON p.Id = k.ParentKategori_Id) AS T " +
                        "WHERE ParentId IS NULL OR (ParentId <> '3004' OR (ParentId = '3004' AND [id] = '3010')) " +
                        "ORDER BY CASE WHEN ParentId IS NULL THEN [order] ELSE (SELECT [order] FROM Kategori WHERE Id = T.ParentId) END, " +
                        "[SortOrder], [Level], [order]";
                    DataTable dth2 = conn.GetDataTable(sql2);
                    rowkeluar = dth2.Rows.Count + row;
                    foreach (DataRow dr2 in dth2.Rows)
                    {
                        worksheet.Cell(row, 1).Value = dr2["Name"].ToString();
                        string? fbold = dr2["SortOrder"].ToString();
                        if (fbold == "1") { worksheet.Cell(row, 1).Style.Font.Bold = true; }
                        else { worksheet.Cell(row, 1).Style.Font.Bold = false; }
                        col = 2;
                        string qtotal = "WITH TanggalProyeksi AS " +
                            "(SELECT DATEADD(DAY, number, CONVERT(date, GETDATE())) AS Tanggal " +
                            "FROM master.dbo.spt_values " +
                            "WHERE type = 'P' " +
                            "AND DATEADD(DAY, number, CONVERT(date, GETDATE())) <= DATEADD(YEAR, 1, CONVERT(date, GETDATE())) " +
                            "AND DATEPART(WEEKDAY, DATEADD(DAY, number, CONVERT(date, GETDATE()))) NOT IN (1, 7)) " +
                            "SELECT CAST(tp.Tanggal AS date) AS Tanggal, pt.Kategori, pt.SubKategori, " +
                            "REPLACE(SUM(pt.Nominal), '.00', '') AS Nominal, " +
                            "(SELECT REPLACE(CAST(SUM(pt2.Nominal) AS varchar), '.00', '') " +
                            "FROM ProTrxFinansial_Log p2 " +
                            "JOIN ProTrxFinansialItem pt2 ON p2.Data_Id = pt2.ProTrxFinansial_Id " +
                            "WHERE p2.TypeTransaksi = '2' " +
                            "AND CAST(p2.TanggalProyeksi AS date) = CAST(tp.Tanggal AS date)) AS TotalKategori " +
                            "FROM TanggalProyeksi tp " +
                            "LEFT JOIN ProTrxFinansial_Log p ON CONVERT(date, p.TanggalProyeksi) = CAST(tp.Tanggal AS date) " +
                            "LEFT JOIN ProTrxFinansialItem pt ON p.Data_Id = pt.ProTrxFinansial_Id " +
                            "WHERE (pt.Kategori IS NULL OR pt.SubKategori IS NULL OR " +
                            "(pt.Kategori = '" + dr2["ParentId"].ToString() + "' AND pt.SubKategori = '" + dr2["SubId"].ToString() + "')) " +
                            "AND (p.TypeTransaksi = '2' OR p.TypeTransaksi IS NULL) " +
                            "GROUP BY CAST(tp.Tanggal AS date), pt.Kategori, pt.SubKategori " +
                            "ORDER BY CAST(tp.Tanggal AS date)";

                        DataTable dttotal = conn.GetDataTable(qtotal);
                        foreach (DataRow drtotal in dttotal.Rows)
                        {
                            if (drtotal["TotalKategori"] != null && !DBNull.Value.Equals(drtotal["TotalKategori"]))
                            {
                                worksheet.Cell(row, col).Value = drtotal["Nominal"].ToString(); 
                                string danamasuk = worksheet.Cell(rowhasil, col).Value.ToString();
                                string? danakeluar = drtotal["TotalKategori"].ToString();
                                decimal decimalA = decimal.TryParse(danamasuk, out decimal decimalValueA) ? decimalValueA : 0;
                                decimal decimalB = decimal.TryParse(danakeluar, out decimal decimalValueB) ? decimalValueB : 0;
                                decimal hasil = decimalB-decimalA;
                                worksheet.Cell(rowkeluar, col).Value = danakeluar;
                                worksheet.Cell(rowkeluar+1, col).Value = hasil;
                            } 
                            col++;      
                        }
                        row++;
                    }
                    Console.WriteLine("Proses Kedua Selesai");
                  
                    worksheet.Cell(row, 1).Value = "Total Dana Keluar";
                    headerRange = worksheet.Range(worksheet.Cell(row, 1), worksheet.Cell(row, col - 1));
                    headerRange.Style.Font.FontColor = XLColor.Black;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                    row++;

                    worksheet.Cell(row, 1).Value = "Net Posisi Cashflow";
                    headerRange = worksheet.Range(worksheet.Cell(row, 1), worksheet.Cell(row, col - 1));
                    headerRange.Style.Font.FontColor = XLColor.Black;
                    headerRange.Style.Fill.BackgroundColor = XLColor.Orange;

                    worksheet.Columns().AdjustToContents();

                    string tanggal = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string filePath = $@"C:\ReportHistoryCashFlow";

                    bool exists = Directory.Exists(filePath);

                    if (!exists) Directory.CreateDirectory(filePath);
                    workbook.SaveAs(filePath + $@"\ReportHistoryCashflow_{tanggal}.xlsx");

                    Console.WriteLine("Data exported to ReportHistoryCashflow.xlsx");

                    Thread.Sleep(nsleep * 86400 * 1000);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        protected override void OnStop()
        {
            try
            {
                if (Worker != null & Worker.IsAlive)
                {
                    Worker.Abort();
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void OnDebug()
        {
            Onstart(null);
        }
    }
}
