using Dapper;
using Microsoft.EntityFrameworkCore;
using ReportHistoryCashflow.Model;
using System.Data;
using System.Data.Common;
using Microsoft.Data.SqlClient;



namespace ReportHistoryCashflow.Data
{
    public class DataContext : DbContext
    {
        public DbSet<Model.Kategori> Kategori { get; set; }
        public DbSet<Model.ProTrxFinansial> ProTrxFinansial { get; set; }
        public DbSet<Model.ProTrxFinansial_Log> ProTrxFinansial_Log { get; set; }
        public DbSet<Model.ProTrxFinansialItem> ProTrxFinansialItem { get; set; }
        public DbSet<CashflowReportItem> CashflowReportItems { get; set; }
        public DataContext(DbContextOptions<DataContext> options) : base(options)
        {
        }

        public virtual List<CashflowReportItem> SUMReportHistoryCashflow(int? kategori, int? subkategori, string type)
        {
            var kategoriParam = new SqlParameter("@kategori", (kategori!= null) ? kategori : (object)DBNull.Value);
            var subkategoriParam = new SqlParameter("@subkategori", (subkategori != null) ? subkategori : (object)DBNull.Value);
            var typeParam = new SqlParameter("@type", type);

            return this.CashflowReportItems
                .FromSqlRaw("EXEC [dbo].[SUMReportHistoryCashflow] @kategori, @subkategori, @type",
                    kategoriParam, subkategoriParam, typeParam)
                .ToList();
        }


        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<CashflowReportItem>().HasNoKey();
        }
    }
}
