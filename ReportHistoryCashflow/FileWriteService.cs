﻿using ClosedXML.Excel;
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
            int rowmasuk = 0;

            try
            {
                while (true)
                {

                    sql conn = new sql();

                    DateTime currentDate = DateTime.Now;
                    DateTime nextYearDate = currentDate.AddYears(1);

                    var workbook = new XLWorkbook();
                    var worksheet = workbook.Worksheets.Add("ReportHistoryCashflow");

                    int col = 1;
                    int row = 4;

                    string sqlcount = "SELECT distinct([Name]), Id FROM Kategori where id in('3025','3003','3006','3004') or ParentKategori_Id in('3025','3003','3006','3004') ORDER BY [Name] DESC";
                    DataTable rc = conn.GetDataTable(sqlcount);
                    int rowhasil = rc.Rows.Count + row;

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
                    Console.WriteLine("tanggal selesai");
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

                    foreach (DataRow dr in dth.Rows)
                    {
                        worksheet.Cell(row, 1).Value = dr["Name"].ToString();
                        string? fbold = dr["SortOrder"].ToString();
                        if (fbold == "1") { worksheet.Cell(row, 1).Style.Font.Bold = true; }
                        else { worksheet.Cell(row, 1).Style.Font.Bold = false; }

                        for (int j = 2; j < col; j++)
                        {
                            DateTime date = DateTime.ParseExact(worksheet.Cell(2, j).Value.ToString(), "dd-MMM-yyyy", CultureInfo.InvariantCulture);
                            string day1 = date.ToString("yyyy-MM-dd" + " 00:00:00.000");
                            string day2 = date.ToString("yyyy-MM-dd" + " 23:59:59.000");

                            string qtotal = "SELECT pt.Kategori, pt.SubKategori, REPLACE(SUM(pt.Nominal), '.00', '') as Nominal," +
                                "(SELECT REPLACE(SUM(pt2.Nominal), '.00', '') FROM ProTrxFinansial_Log p2 " +
                                "JOIN ProTrxFinansialItem pt2 ON p2.Data_Id = pt2.ProTrxFinansial_Id " +
                                "WHERE p2.TypeTransaksi = '1' AND p2.TanggalProyeksi " +
                                "BETWEEN '" + day1 + "' AND '" + day2 + "') as TotalKategori " +
                                "FROM ProTrxFinansial_Log p JOIN  ProTrxFinansialItem pt ON p.Data_Id = pt.ProTrxFinansial_Id " +
                                "WHERE p.TypeTransaksi = '1' AND p.TanggalProyeksi BETWEEN '" + day1 + "' AND '" + day2 + "' AND pt.Kategori = '" + dr["ParentId"].ToString() + "' AND pt.SubKategori = '" + dr["SubId"].ToString() + "'" +
                                " GROUP BY pt.Kategori, " +
                                "pt.SubKategori " +
                                "HAVING (pt.Kategori IS NOT NULL OR pt.SubKategori IS NOT NULL) AND pt.SubKategori IS NOT NULL";

                            DataTable dttotal = conn.GetDataTable(qtotal);

                            foreach (DataRow drtotal in dttotal.Rows)
                            {
                                worksheet.Cell(row, j).Value = drtotal["Nominal"].ToString();
                                worksheet.Cell(rowhasil, j).Value = drtotal["TotalKategori"].ToString();
                            }
                        }
                        row++;
                    }
                    Console.WriteLine("Proses 1 selesai");
                    var headerRange = worksheet.Range(worksheet.Cell(3, 1), worksheet.Cell(3, col - 1));
                    headerRange.Merge();
                    headerRange.Value = "Dana Masuk ";
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Font.FontColor = XLColor.Black;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightPink;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    rowmasuk = row;
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
                    int rowkeluar = dth2.Rows.Count + row;
                   
                    foreach (DataRow dr2 in dth2.Rows)
                    {
                        worksheet.Cell(row, 1).Value = dr2["Name"].ToString();
                        string? fbold = dr2["SortOrder"].ToString();
                        if (fbold == "1") { worksheet.Cell(row, 1).Style.Font.Bold = true; }
                        else { worksheet.Cell(row, 1).Style.Font.Bold = false; }

                        for (int j = 2; j < col; j++)
                        {
                            DateTime date = DateTime.ParseExact(worksheet.Cell(2, j).Value.ToString(), "dd-MMM-yyyy", CultureInfo.InvariantCulture);
                            string day1 = date.ToString("yyyy-MM-dd" + " 00:00:00.000");
                            string day2 = date.ToString("yyyy-MM-dd" + " 23:59:59.000");
                            var hasilmasuk = worksheet.Cell(rowmasuk, j).Value;
                            string qtotal = $"SELECT pt.Kategori, pt.SubKategori, REPLACE(SUM(pt.Nominal), '.00', '') as Nominal," +
                                "(SELECT REPLACE(SUM(pt2.Nominal), '.00', '') FROM ProTrxFinansial_Log p2 " +
                                "JOIN ProTrxFinansialItem pt2 ON p2.Data_Id = pt2.ProTrxFinansial_Id " +
                                "WHERE p2.TypeTransaksi = '2' AND p2.TanggalProyeksi " +
                                "BETWEEN '"+ day1 + "' AND '"+ day2+"') as TotalKategori, " +
                                "(SELECT REPLACE(ISNULL(SUM(pt2.Nominal), 0)-'"+hasilmasuk+"', '.00', '') FROM ProTrxFinansial_Log p2 " +
                                "JOIN ProTrxFinansialItem pt2 ON p2.Data_Id = pt2.ProTrxFinansial_Id " +
                                "WHERE p2.TypeTransaksi = '2' AND p2.TanggalProyeksi " +
                                "BETWEEN '"+ day1 + "' AND '"+ day2 + "') as Cashflow " +
                                "FROM ProTrxFinansial_Log p JOIN  ProTrxFinansialItem pt ON p.Data_Id = pt.ProTrxFinansial_Id " +
                                "WHERE p.TypeTransaksi = '2' AND p.TanggalProyeksi BETWEEN '"+day1+"' AND '"+day2+ "' AND pt.Kategori = '"+ dr2["ParentId"].ToString() + "' AND pt.SubKategori = '"+ dr2["SubId"].ToString() + "'" +
                                " GROUP BY pt.Kategori, " +
                                "pt.SubKategori " +
                                "HAVING (pt.Kategori IS NOT NULL OR pt.SubKategori IS NOT NULL) AND pt.SubKategori IS NOT NULL";

                            DataTable dttotal = conn.GetDataTable(qtotal);

                            foreach (DataRow drtotal in dttotal.Rows)
                            {
                                worksheet.Cell(row, j).Value = drtotal["Nominal"].ToString();
                                worksheet.Cell(rowkeluar, j).Value = drtotal["TotalKategori"].ToString();
                                worksheet.Cell(rowkeluar+1, j).Value = drtotal["Cashflow"].ToString();
                            }
                        }
                        row++;
                    }
                    Console.WriteLine("proses 2 selesai");
                    worksheet.Cell(row, 1).Value = "Total Dana Keluar";
                    headerRange = worksheet.Range(worksheet.Cell(row, 1), worksheet.Cell(row, col - 1));
                    headerRange.Style.Font.FontColor = XLColor.Black;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                    row++;

                    worksheet.Cell(row, 1).Value = "Net Posisi Cashflow";
                    headerRange = worksheet.Range(worksheet.Cell(row, 1), worksheet.Cell(row, col - 1));
                    headerRange.Style.Font.FontColor = XLColor.Black;
                    headerRange.Style.Fill.BackgroundColor = XLColor.Orange;

                    //string sql = "SELECT distinct([Name]), Id FROM Kategori where id in('3025','3003','3006','3004') ORDER BY [Name] DESC";
                    //DataTable dth = conn.GetDataTable(sql);
                    //string sqlcount = "SELECT distinct([Name]), Id FROM Kategori where id in('3025','3003','3006','3004') or ParentKategori_Id in('3025','3003','3006','3004') ORDER BY [Name] DESC";
                    //DataTable rc = conn.GetDataTable(sqlcount);
                    //int rowhasil = rc.Rows.Count + row;
                    //foreach (DataRow dr in dth.Rows)
                    //{
                    //    worksheet.Cell(row, 1).Value = dr["Name"].ToString();
                    //    worksheet.Cell(row, 1).Style.Font.Bold = true;
                    //    row++;

                    //    string query = "SELECT [Name], [id] FROM Kategori WHERE [TYPE] = '2' AND ParentKategori_Id = " + dr["Id"].ToString() + "";
                    //    DataTable dtd = conn.GetDataTable(query);

                    //    foreach (DataRow drd in dtd.Rows)
                    //    {
                    //        worksheet.Cell(row, 1).Value = "      " + drd["Name"].ToString();
                    //        for (int j = 2; j < col; j++)
                    //        {
                    //            DateTime date = DateTime.ParseExact(worksheet.Cell(2, j).Value.ToString(), "dd-MMM-yyyy", CultureInfo.InvariantCulture);
                    //            string day1 = date.ToString("yyyy-MM-dd" + " 00:00:00.000");
                    //            string day2 = date.ToString("yyyy-MM-dd" + " 23:59:59.000");

                    //        string rs = "select " +
                    //            "replace(pt.Nominal,'.00','') as nominal " +
                    //            "from ProTrxFinansial_Log p " +
                    //            "join ProTrxFinansialItem pt on p.Data_Id = pt.ProTrxFinansial_Id " +
                    //            "where p.TypeTransaksi ='1' and pt.Kategori ='" + dr["Id"].ToString() + "' and pt.SubKategori='" +
                    //            drd["id"].ToString() + "' and TanggalProyeksi between '" + day1 + "' and '" + day2 + "'";
                    //        DataTable dtv = conn.GetDataTable(rs);
                    //        if (dtv != null)
                    //        {
                    //            foreach (DataRow drv in dtv.Rows)
                    //            {
                    //                worksheet.Cell(row, j).Value = drv["Nominal"].ToString();
                    //            }
                    //        }
                    //        string rss = "select replace(ISNULL(SUM(pt.Nominal), 0), '.00','') as Nominal" +
                    //           " from ProTrxFinansial_Log p" +
                    //           " join ProTrxFinansialItem pt on p.Data_Id = pt.ProTrxFinansial_Id " +
                    //           "where p.TypeTransaksi='1' and p.TanggalProyeksi between '" + day1 + "' and '" + day2 + "'";
                    //        DataTable dts = conn.GetDataTable(rss);
                    //        if (dts != null)
                    //        {
                    //            foreach (DataRow drs in dts.Rows)
                    //            {
                    //                worksheet.Cell(rowhasil, j).Value = drs["Nominal"].ToString();
                    //            }
                    //        }
                    //    }
                    //    row++;
                    //    }
                    //}

                    //rowmasuk = row;
                    //worksheet.Cell(row, 1).Value = "Total Dana Masuk";
                    //headerRange = worksheet.Range(worksheet.Cell(row, 1), worksheet.Cell(row, col - 1));
                    //headerRange.Style.Font.FontColor = XLColor.Black;
                    //headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                    //row++;

                    //headerRange = worksheet.Range(worksheet.Cell(row, 1), worksheet.Cell(row, col - 1));
                    //headerRange.Merge();
                    //headerRange.Value = "Dana Keluar";
                    //headerRange.Style.Font.Bold = true;
                    //headerRange.Style.Font.FontColor = XLColor.Black;
                    //headerRange.Style.Fill.BackgroundColor = XLColor.LightPink;
                    //headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    //headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    //row++;

                    //string sql2 = "SELECT distinct([Name]), Id FROM Kategori where id in('3025','3003','3006','8') ORDER BY [Name] DESC";
                    //DataTable dth2 = conn.GetDataTable(sql2);

                    //foreach (DataRow dr2 in dth2.Rows)
                    //{
                    //    worksheet.Cell(row, 1).Value = dr2["Name"].ToString();
                    //    worksheet.Cell(row, 1).Style.Font.Bold = true;
                    //    row++;

                    //    string query2 = "SELECT [Name], [id] FROM Kategori WHERE [TYPE] = '2' AND ParentKategori_Id = " + dr2["Id"].ToString() + "";
                    //    DataTable dtd2 = conn.GetDataTable(query2);

                    //    foreach (DataRow drd2 in dtd2.Rows)
                    //    {
                    //        worksheet.Cell(row, 1).Value = "      " + drd2["Name"].ToString();
                    //        for (int j = 2; j < col; j++)
                    //        {
                    //            DateTime date = DateTime.ParseExact(worksheet.Cell(2, j).Value.ToString(), "dd-MMM-yyyy", CultureInfo.InvariantCulture);
                    //            string day1 = date.ToString("yyyy-MM-dd" + " 00:00:00.000");
                    //            string day2 = date.ToString("yyyy-MM-dd" + " 23:59:59.000");
                    //            string rs = "select " +
                    //                "replace(SUM(pt.Nominal),'.00','') as nominal " +
                    //                "from ProTrxFinansial_Log p " +
                    //                "join ProTrxFinansialItem pt on p.Data_Id = pt.ProTrxFinansial_Id " +
                    //                "where p.TypeTransaksi ='2' and pt.Kategori ='" + dr2["Id"].ToString() + "' and pt.SubKategori='" +
                    //                drd2["id"].ToString() + "' and TanggalProyeksi between '" + day1 + "' and '" + day2 + "'";
                    //            DataTable dtv = conn.GetDataTable(rs);
                    //            if (dtv != null)
                    //            {
                    //                foreach (DataRow drv in dtv.Rows)
                    //                {
                    //                    worksheet.Cell(row, j).Value = drv["Nominal"].ToString();
                    //                }
                    //            }
                    //        }
                    //        row++;
                    //    }
                    //}

                    //for (int j = 2; j < col; j++)
                    //{
                    //    DateTime date = DateTime.ParseExact(worksheet.Cell(2, j).Value.ToString(), "dd-MMM-yyyy", CultureInfo.InvariantCulture);
                    //    string day1 = date.ToString("yyyy-MM-dd" + " 00:00:00.000");
                    //    string day2 = date.ToString("yyyy-MM-dd" + " 23:59:59.000");

                    //    var hasilmasuk = (string)worksheet.Cell(rowmasuk, j).Value != "" ? worksheet.Cell(rowmasuk, j).Value : 0;
                    //    string rss = "select replace(ISNULL(SUM(pt.Nominal),0), '.00','') as Nominal," +
                    //        "replace(ISNULL(SUM(pt.Nominal) - " + hasilmasuk + ",0), '.00','') as Total" +
                    //        " from ProTrxFinansial_Log p" +
                    //        " join ProTrxFinansialItem pt on p.Data_Id = pt.ProTrxFinansial_Id " +
                    //        "where p.TypeTransaksi='2' and p.TanggalProyeksi between '" + day1 + "' and '" + day2 + "'";
                    //    DataTable dts = conn.GetDataTable(rss);
                    //    if (dts != null)
                    //    {
                    //        foreach (DataRow drs in dts.Rows)
                    //        {
                    //            worksheet.Cell(row, j).Value = drs["Nominal"].ToString();
                    //            worksheet.Cell(row + 1, j).Value = drs["Total"].ToString();
                    //        }
                    //    }
                    //}
                    //worksheet.Cell(row, 1).Value = "Total Dana Keluar";
                    //headerRange = worksheet.Range(worksheet.Cell(row, 1), worksheet.Cell(row, col - 1));
                    //headerRange.Style.Font.FontColor = XLColor.Black;
                    //headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                    //row++;

                    //worksheet.Cell(row, 1).Value = "Net Posisi Cashflow";
                    //headerRange = worksheet.Range(worksheet.Cell(row, 1), worksheet.Cell(row, col - 1));
                    //headerRange.Style.Font.FontColor = XLColor.Black;
                    //headerRange.Style.Fill.BackgroundColor = XLColor.Orange;

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
