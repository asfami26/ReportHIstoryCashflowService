using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportHistoryCashflow.Model
{
    public class CashflowReportItem
    {
        public DateTime Tanggal { get; set; }
        public int Kategori { get; set; }
        public int SubKategori { get; set; }
        public string? Nominal { get; set; }
        public string? TotalKategori { get; set; }
    }
}
