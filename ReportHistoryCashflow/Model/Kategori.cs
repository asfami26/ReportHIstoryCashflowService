using System;

namespace ReportHistoryCashflow.Model
{
    public class Kategori
    {
        public int Id { get; set; }
        public int Type { get; set; }
        public string Name { get; set; }
        public int? Order { get; set; }
        public int? ParentKategori_Id { get; set; }
        public DateTime? CreatedTime { get; set; }
        public DateTime? UpdatedTime { get; set; }
        public int? CreatedBy_Id { get; set; }
        public int? UpdatedBy_Id { get; set; }
        public int Level { get; set; }
        public int? Value { get; set; }

        public class KategoriResult
        {
            public string Name { get; set; }
            public int Id { get; set; }
            public int SortOrder { get; set; }
            public int? ParentId { get; set; }
            public int? SubId { get; set; }
            public int Level { get; set; }
        }

    }
}
