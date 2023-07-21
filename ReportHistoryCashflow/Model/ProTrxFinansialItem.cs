using System;

namespace ReportHistoryCashflow.Model
{
    public class ProTrxFinansialItem
    {
        public int Id { get; set; }
        public int ProTrxFinansial_Id { get; set; }
        public int CCY { get; set; }
        public string? CCYName { get; set; }
        public decimal? Rate { get; set; }
        public decimal Nominal { get; set; }
        public string? Catatan { get; set; }
        public int Kategori { get; set; }
        public string? KategoriName { get; set; }
        public string? NamaNasabah { get; set; }
        public string? KodeBankTerkait { get; set; }
        public string? NamaBankTerkait { get; set; }
        public int SubKategori { get; set; }
        public string? SubKategoriName { get; set; }
        public bool? Verify { get; set; }
        public int? TujuanTransaksi { get; set; }
        public string? TujuanTransaksiName { get; set; }
        public int? KriteriaNasabah { get; set; }
        public string? KriteriaNasabahName { get; set; }
        public int? StatusKredit { get; set; }
        public string? StatusKreditName { get; set; }

        // Optional: You can add navigation properties here if needed
    }
}

