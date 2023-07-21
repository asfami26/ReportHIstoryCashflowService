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
                            int? parentId = item.ParentId; 
                            int? subId = item.SubId; 
                            string type = "2";

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
            stopEvent.Set();
            
            if (!workerThread.Join(TimeSpan.FromSeconds(10)))
            {
                
            }

            stopEvent.Dispose();
        }

        public void OnDebug()
        {
            Onstart(null!);
        }
    }
}
