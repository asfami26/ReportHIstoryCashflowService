﻿using System;

namespace ReportHistoryCashflow.Model
{
    public class ProTrxFinansial_Log
    {
        public int Id { get; set; }
        public int Data_ID { get; set; }
        public int UnitId { get; set; }
        public string? UnitName { get; set; }
        public string? UnitCode { get; set; }
        public int? RegionId { get; set; }
        public string? RegionName { get; set; }
        public string? RegionCode { get; set; }
        public DateTime TanggalProyeksi { get; set; }
        public int ItemCount { get; set; }
        public int TypeTransaksi { get; set; }
        public string? TypeTransaksiName { get; set; }
        public int ActionBy { get; set; }
        public string? Action { get; set; }
        public DateTime? ActionTime { get; set; }
        public int? Status { get; set; }
        public string? StatusName { get; set; }
        public string? Approver { get; set; }
        public string? Keterangan_Approval { get; set; }
        public int? Flow { get; set; }
        public string? FlowName { get; set; }
        public string? UnitIdApprover { get; set; }
        public string? UnitNameApprover { get; set; }
        public string? Role { get; set; }
        public int? ModifyRequest { get; set; }

        // Optional: You can add navigation properties here if needed
    }
}
