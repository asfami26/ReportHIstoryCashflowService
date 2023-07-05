using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml.Wordprocessing;
using ReportHistoryCashflow.Class;
using System.Collections;

namespace ReportHistoryCashflow
{
    public class FileWriteService :ServiceBase
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

            try
            {
                while (true)
                {
                    sql conn = new sql();
                    string sql = "SELECT distinct([Name]), Id FROM Kategori where id in('3025','3003','3006','3004') ORDER BY [Name] DESC";
                    DataTable dth = conn.GetDataTable(sql);


                    DateTime currentDate = DateTime.Now;
                    DateTime nextYearDate = currentDate.AddYears(1);

                    var workbook = new XLWorkbook();
                    var worksheet = workbook.Worksheets.Add("ReportHistoryCashflow");

                    int col = 1;
                    var headerCellA2 = worksheet.Cell(2, col);
                    headerCellA2.Value = "Keterangan";
                    headerCellA2.Style.Fill.BackgroundColor = XLColor.TwilightLavender;
                    headerCellA2.Style.Font.FontColor = XLColor.White;

                    col++;
                    while (currentDate <= nextYearDate)
                    {
                        var headerCell = worksheet.Cell(2, col);
                        headerCell.Value = currentDate.ToString("dd-MMM-yyyy");
                        headerCell.Style.Fill.BackgroundColor = XLColor.TwilightLavender;
                        headerCell.Style.Font.FontColor = XLColor.White;
                        currentDate = currentDate.AddDays(1);
                        col++;
                    }

                    var headerRange = worksheet.Range(worksheet.Cell(3, 1), worksheet.Cell(3, col - 1));
                    headerRange.Merge();
                    headerRange.Value = "Dana Masuk ";
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Font.FontColor = XLColor.Black;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightPink;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    int row = 4;

                    foreach (DataRow dr in dth.Rows)
                    {
                        worksheet.Cell(row, 1).Value = dr["Name"].ToString();
                        worksheet.Cell(row, 1).Style.Font.Bold = true;
                        row++;

                        string query = "SELECT [Name] FROM Kategori WHERE [TYPE] = '2' AND ParentKategori_Id = "+ dr["Id"].ToString() + "";
                        DataTable dtd = conn.GetDataTable(query);

                        foreach (DataRow drd in dtd.Rows)
                        {
                            worksheet.Cell(row, 1).Value = "      " + drd["Name"].ToString();
                            row++;
                        }
                    }

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


                    string sql2 = "SELECT distinct([Name]), Id FROM Kategori where id in('3025','3003','3006','8') ORDER BY [Name] DESC";
                    DataTable dth2 = conn.GetDataTable(sql2);

                    foreach (DataRow dr2 in dth2.Rows)
                    {
                        worksheet.Cell(row, 1).Value = dr2["Name"].ToString();
                        worksheet.Cell(row, 1).Style.Font.Bold = true;
                        row++;

                        string query2 = "SELECT [Name] FROM Kategori WHERE [TYPE] = '2' AND ParentKategori_Id = " + dr2["Id"].ToString() + "";
                        DataTable dtd2 = conn.GetDataTable(query2);

                        foreach (DataRow drd2 in dtd2.Rows)
                        {
                            worksheet.Cell(row, 1).Value = "      " + drd2["Name"].ToString();
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
                    string filePath = $@"D:\ReportHistoryCashflow_{tanggal}.xlsx";
                    workbook.SaveAs(filePath);

                    Console.WriteLine("Dates exported to ReportHistoryCashflow.xlsx");

                    Thread.Sleep(nsleep * 60 * 1000);
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
