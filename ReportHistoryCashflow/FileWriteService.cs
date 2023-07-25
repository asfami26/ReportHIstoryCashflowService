using ClosedXML.Excel;
using System.Data;
using ReportHistoryCashflow.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using static ReportHistoryCashflow.Model.Kategori;
using Microsoft.Extensions.Hosting;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using System.IO;
using System.Linq;
using System;
using System.Collections.Generic;

namespace ReportHistoryCashflow
{
    public class FileWriteService : IHostedService
    {
        private CancellationTokenSource? _cts;
        private readonly IHostApplicationLifetime _appLifetime;

        public FileWriteService(IHostApplicationLifetime appLifetime)
        {
            _appLifetime = appLifetime;
            _cts = new CancellationTokenSource();
        }

        public Task StartAsync(CancellationToken cancellationToken)
        {
            _cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            Task.Run(ExecuteAsync, _cts.Token);
            return Task.CompletedTask;
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            _cts?.Cancel();
            return Task.CompletedTask;
        }

        private async Task ExecuteAsync()
        {
            try
            {
                string basePath = Directory.GetCurrentDirectory();

                IConfigurationRoot configuration = new ConfigurationBuilder()
                    .SetBasePath(basePath)
                    .AddJsonFile("appsettings.json")
                    .Build();


                var optionsBuilder = new DbContextOptionsBuilder<DataContext>();
                optionsBuilder.UseSqlServer(configuration.GetConnectionString("DbConnection"));

                var options = optionsBuilder.Options;

                int rowhasil = 0;
                int col = 0;
                int row = 0;
                int rowkeluar = 0;
                while (_cts != null && !_cts.Token.IsCancellationRequested)
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
                        var result = KategoriQuery.GetKategoriResults(new int[] { 3025, 3003, 3006, 3004 }, 3008, dbContext);
                        row = 4;
                        rowhasil = result.Count() + row;
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
                        var result = KategoriQuery.GetKategoriResults(new int[] { 3025, 3003, 3006, 3004 }, 3010, dbContext);
                        rowkeluar = result.Count() + row;
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
                                    string danakeluar = val?.TotalKategori?.ToString() ?? "";
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

                    await Task.Delay(1000, _cts.Token); // Contoh: menunda selama 1 detik
                    StopHost();
                }
            }
            catch (TaskCanceledException)
            {
                // Tugas dibatalkan, tidak perlu melakukan apa-apa
            }
            catch (Exception ex)
            {
                // Tangani kesalahan lainnya sesuai kebutuhan Anda
                Console.WriteLine("Terjadi kesalahan: " + ex.Message);
            }
        }

        public void StopHost()
        {
            _cts?.Cancel();
            _appLifetime.StopApplication();
        }


        public static class KategoriQuery
        {
            public static List<KategoriResult> GetKategoriResults(int[] param1, int param2, DataContext dbContext)
            {
                var query = (
                    from k1 in dbContext.Kategori
                    where param1.Contains(k1.Id)
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
                        where param1.Contains(k3.Id)
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
                .Where(t => t.ParentId == null || (t.ParentId != 3004 || (t.ParentId == 3004 && t.Id == param2)));

                var orderedQuery = query.ToList()
                                   .OrderBy(t => t.ParentId == null ? dbContext.Kategori.FirstOrDefault(k => k.Id == t.Id)?.Order : dbContext.Kategori.FirstOrDefault(k => k.Id == t.ParentId)?.Order)
                                   .ThenBy(t => t.SortOrder)
                                   .ThenBy(t => t.Level)
                                   .ThenBy(t => t.Id);

                return orderedQuery.ToList();
            }
        }
    }
}
