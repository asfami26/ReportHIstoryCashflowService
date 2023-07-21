using ClosedXML.Excel;
using System.Data;
using System.ServiceProcess;
using ReportHistoryCashflow.Class;
using ReportHistoryCashflow.Data;
using System.Linq;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using DocumentFormat.OpenXml.InkML;
using Microsoft.Extensions.DependencyInjection;
using System.Web.Services.Description;
using System;
using ReportHistoryCashflow.Model;
using static ReportHistoryCashflow.Model.Kategori;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data.SqlClient;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Vml;

namespace ReportHistoryCashflow
{
    public class FileWriteService : ServiceBase
    {

        private Thread workerThread;
        private ManualResetEvent stopEvent;
        public FileWriteService()
        {
            ServiceName = "ReportHistoryCashflow";
        }

        protected void Onstart(string[] args)
        {
            stopEvent = new ManualResetEvent(false);
            workerThread = new Thread(Working);
            workerThread.Start();
        }

        public void Working()
        {

            string basePath = Directory.GetCurrentDirectory();

            IConfigurationRoot configuration = new ConfigurationBuilder()
                .SetBasePath(basePath)
                .AddJsonFile("appsettings.json")
                .Build();


            var optionsBuilder = new DbContextOptionsBuilder<DataContext>();
            optionsBuilder.UseSqlServer(configuration.GetConnectionString("DbConnection"));

            var options = optionsBuilder.Options;

            int nsleep = 1;
            int rowhasil = 0;
            int col = 0;
            int row = 0;
            int rowkeluar = 0;
            try
            {
                while (!stopEvent.WaitOne(0))
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

                    using (var dbContext = new DataContext(options))
                    {
                        var query = (
                             from k1 in dbContext.Kategori
                             where new[] { 3025, 3003, 3006, 3004 }.Contains(k1.Id)
                             select new KategoriResult
                             {
                                 Name = k1.Name,
                                 Id = k1.Id,
                                 SortOrder = 1,
                                 ParentId = k1.ParentKategori_Id,
                                 SubId = null,
                                 Level = 0
                             }
                         )
                         .Concat(
                             from k2 in dbContext.Kategori
                             join p in (
                                 from k3 in dbContext.Kategori
                                 where new[] { 3025, 3003, 3006, 3004 }.Contains(k3.Id)
                                 select new
                                 {
                                     Id = k3.Id,
                                     Level = 0,
                                     ParentId = k3.Id,
                                     Order = k3.Order
                                 }
                             ) on k2.ParentKategori_Id equals p.Id
                             select new KategoriResult
                             {
                                 Name = k2.Id == 3004 && k2.Name == "Remis" ? "      " + k2.Name : "      " + k2.Name,
                                 Id = k2.Id,
                                 SortOrder = p.Level + 1,
                                 ParentId = k2.ParentKategori_Id,
                                 SubId = k2.Id,
                                 Level = p.Level + 1
                             }
                         )
                         .Where(t => t.ParentId == null || (t.ParentId != 3004 || (t.ParentId == 3004 && t.Id == 3008)));

                        var orderedQuery = query.ToList()
                                           .OrderBy(t => t.ParentId == null ? dbContext.Kategori.FirstOrDefault(k => k.Id == t.Id)?.Order : dbContext.Kategori.FirstOrDefault(k => k.Id == t.ParentId)?.Order)
                                           .ThenBy(t => t.SortOrder)
                                           .ThenBy(t => t.Level)
                                           .ThenBy(t => t.Id);

                        var result = orderedQuery.ToList();

                        row = 4;
                        rowhasil = orderedQuery.Count() + row;
                        foreach (var item in result)
                        {
                            worksheet.Cell(row, 1).Value = item.Name;
                            if (item.SubId == null) { worksheet.Cell(row, 1).Style.Font.Bold = true; }
                            else { worksheet.Cell(row, 1).Style.Font.Bold = false; }

                            col = 2;

                            int? parentId = item.ParentId ?? 0;
                            int? subId = item.SubId ?? 0;
                            string type = "1";

                            var results = dbContext.SUMReportHistoryCashflow(parentId, subId, type);

                            foreach (var val in results)
                            {
                                var nominal = val.Nominal;
                                var totalKategori = val.TotalKategori;

                                worksheet.Cell(row, col).Value = nominal;
                                if (totalKategori != "0")
                                {
                                    worksheet.Cell(rowhasil, col).Value = totalKategori;
                                }
                                col++;
                            }
                            row++;
                        }
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

                    //string sql2 = "SELECT [Name], [Id], [SortOrder], [ParentId], [SubId] " +
                    //    "FROM (SELECT Id,[Name],0 AS [Level],ParentKategori_Id AS ParentId, " +
                    //    "[order],1 AS [SortOrder],NULL AS [SubId] " +
                    //    "FROM Kategori " +
                    //    "WHERE Id IN ('3025', '3003', '3006', '3004') " +
                    //    "UNION ALL " +
                    //    "SELECT k.Id, CASE WHEN k.Id = '3004' AND k.[Name] = 'Supply' THEN '      ' + k.[Name] " +
                    //    "ELSE '      ' + k.[Name] END, p.[Level] + 1 AS [Level], k.ParentKategori_Id AS ParentId, k.[order], " +
                    //    "2 AS [SortOrder], k.Id AS [SubId] FROM Kategori k INNER JOIN (SELECT Id, 0 AS [Level], Id AS ParentId, [order] " +
                    //    "FROM Kategori " +
                    //    "WHERE Id IN ('3025', '3003', '3006', '3004')) p ON p.Id = k.ParentKategori_Id) AS T " +
                    //    "WHERE ParentId IS NULL OR (ParentId <> '3004' OR (ParentId = '3004' AND [id] = '3010')) " +
                    //    "ORDER BY CASE WHEN ParentId IS NULL THEN [order] ELSE (SELECT [order] FROM Kategori WHERE Id = T.ParentId) END, " +
                    //    "[SortOrder], [Level], [order]";
                    //DataTable dth2 = conn.GetDataTable(sql2);
                    //rowkeluar = dth2.Rows.Count + row;
                    //foreach (DataRow dr2 in dth2.Rows)
                    //{
                    //    worksheet.Cell(row, 1).Value = dr2["Name"].ToString();
                    //    string? fbold = dr2["SortOrder"].ToString();
                    //    if (fbold == "1") { worksheet.Cell(row, 1).Style.Font.Bold = true; }
                    //    else { worksheet.Cell(row, 1).Style.Font.Bold = false; }
                    //    col = 2;
                    //    string qtotal = "WITH TanggalProyeksi AS " +
                    //        "(SELECT DATEADD(DAY, number, CONVERT(date, GETDATE())) AS Tanggal " +
                    //        "FROM master.dbo.spt_values " +
                    //        "WHERE type = 'P' " +
                    //        "AND DATEADD(DAY, number, CONVERT(date, GETDATE())) <= DATEADD(YEAR, 1, CONVERT(date, GETDATE())) " +
                    //        "AND DATEPART(WEEKDAY, DATEADD(DAY, number, CONVERT(date, GETDATE()))) NOT IN (1, 7)) " +
                    //        "SELECT CAST(tp.Tanggal AS date) AS Tanggal, pt.Kategori, pt.SubKategori, " +
                    //        "REPLACE(SUM(pt.Nominal), '.00', '') AS Nominal, " +
                    //        "(SELECT REPLACE(CAST(SUM(pt2.Nominal) AS varchar), '.00', '') " +
                    //        "FROM ProTrxFinansial_Log p2 " +
                    //        "JOIN ProTrxFinansialItem pt2 ON p2.Data_Id = pt2.ProTrxFinansial_Id " +
                    //        "WHERE p2.TypeTransaksi = '2' " +
                    //        "AND CAST(p2.TanggalProyeksi AS date) = CAST(tp.Tanggal AS date)) AS TotalKategori " +
                    //        "FROM TanggalProyeksi tp " +
                    //        "LEFT JOIN ProTrxFinansial_Log p ON CONVERT(date, p.TanggalProyeksi) = CAST(tp.Tanggal AS date) " +
                    //        "LEFT JOIN ProTrxFinansialItem pt ON p.Data_Id = pt.ProTrxFinansial_Id " +
                    //        "WHERE (pt.Kategori IS NULL OR pt.SubKategori IS NULL OR " +
                    //        "(pt.Kategori = '" + dr2["ParentId"].ToString() + "' AND pt.SubKategori = '" + dr2["SubId"].ToString() + "')) " +
                    //        "AND (p.TypeTransaksi = '2' OR p.TypeTransaksi IS NULL) " +
                    //        "GROUP BY CAST(tp.Tanggal AS date), pt.Kategori, pt.SubKategori " +
                    //        "ORDER BY CAST(tp.Tanggal AS date)";

                    //    DataTable dttotal = conn.GetDataTable(qtotal);
                    //    foreach (DataRow drtotal in dttotal.Rows)
                    //    {
                    //        if (drtotal["TotalKategori"] != null && !DBNull.Value.Equals(drtotal["TotalKategori"]))
                    //        {
                    //            worksheet.Cell(row, col).Value = drtotal["Nominal"].ToString();
                    //            string danamasuk = worksheet.Cell(rowhasil, col).Value.ToString();
                    //            string? danakeluar = drtotal["TotalKategori"].ToString();
                    //            decimal decimalA = decimal.TryParse(danamasuk, out decimal decimalValueA) ? decimalValueA : 0;
                    //            decimal decimalB = decimal.TryParse(danakeluar, out decimal decimalValueB) ? decimalValueB : 0;
                    //            decimal hasil = decimalB - decimalA;
                    //            worksheet.Cell(rowkeluar, col).Value = danakeluar;
                    //            worksheet.Cell(rowkeluar + 1, col).Value = hasil;
                    //        }
                    //        col++;
                    //    }
                    //    row++;
                    //}

                    using (var dbContext = new DataContext(options))
                    {
                        var query = (
                             from k1 in dbContext.Kategori
                             where new[] { 3025, 3003, 3006, 3004 }.Contains(k1.Id)
                             select new KategoriResult
                             {
                                 Name = k1.Name,
                                 Id = k1.Id,
                                 SortOrder = 1,
                                 ParentId = k1.ParentKategori_Id,
                                 SubId = null,
                                 Level = 0
                             }
                         )
                         .Concat(
                             from k2 in dbContext.Kategori
                             join p in (
                                 from k3 in dbContext.Kategori
                                 where new[] { 3025, 3003, 3006, 3004 }.Contains(k3.Id)
                                 select new
                                 {
                                     Id = k3.Id,
                                     Level = 0,
                                     ParentId = k3.Id,
                                     Order = k3.Order
                                 }
                             ) on k2.ParentKategori_Id equals p.Id
                             select new KategoriResult
                             {
                                 Name = k2.Id == 3004 && k2.Name == "Remis" ? "      " + k2.Name : "      " + k2.Name,
                                 Id = k2.Id,
                                 SortOrder = p.Level + 1,
                                 ParentId = k2.ParentKategori_Id,
                                 SubId = k2.Id,
                                 Level = p.Level + 1
                             }
                         )
                         .Where(t => t.ParentId == null || (t.ParentId != 3004 || (t.ParentId == 3004 && t.Id == 3010)));

                        var orderedQuery = query.ToList()
                                           .OrderBy(t => t.ParentId == null ? dbContext.Kategori.FirstOrDefault(k => k.Id == t.Id)?.Order : dbContext.Kategori.FirstOrDefault(k => k.Id == t.ParentId)?.Order)
                                           .ThenBy(t => t.SortOrder)
                                           .ThenBy(t => t.Level)
                                           .ThenBy(t => t.Id);

                        var result = orderedQuery.ToList();
                        rowkeluar = orderedQuery.Count() + row;
                        foreach (var item in result)
                        {
                            worksheet.Cell(row, 1).Value = item.Name;
                            if (item.SubId == null) { worksheet.Cell(row, 1).Style.Font.Bold = true; }
                            else { worksheet.Cell(row, 1).Style.Font.Bold = false; }
                            col = 2;
                            int? parentId = item.ParentId; // Isi dengan nilai yang sesuai
                            int? subId = item.SubId; // Isi dengan nilai yang sesuai
                            string type = "2"; // Isi dengan nilai yang sesuai

                            //query total
                            var results = dbContext.SUMReportHistoryCashflow(parentId, subId, type);

                            foreach (var val in results)
                            {

                                var nominal = val.Nominal;
                                worksheet.Cell(row, col).Value = nominal;
                                if (val.TotalKategori != "0")
                                {
                                    string danamasuk = worksheet.Cell(rowhasil, col).Value.ToString();
                                    string danakeluar = val.TotalKategori.ToString();
                                    decimal decimalA = decimal.TryParse(danamasuk, out decimal decimalValueA) ? decimalValueA : 0;
                                    decimal decimalB = decimal.TryParse(danakeluar, out decimal decimalValueB) ? decimalValueB : 0;
                                    decimal hasil = decimalB - decimalA;
                                    worksheet.Cell(rowkeluar, col).Value = danakeluar;
                                    worksheet.Cell(rowkeluar + 1, col).Value = hasil;
                                }
                                col++;
                            }
                            row++;
                        }
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
            // Set event untuk memberhentikan thread
            stopEvent.Set();

            // Tunggu hingga thread berhenti dengan timeout
            if (!workerThread.Join(TimeSpan.FromSeconds(10)))
            {
                // Jika thread tidak berhenti dalam waktu yang ditentukan, lakukan tindakan yang sesuai
            }

            // Bersihkan event
            stopEvent.Dispose();
        }

        public void OnDebug()
        {
            Onstart(null!);
        }
    }
}
