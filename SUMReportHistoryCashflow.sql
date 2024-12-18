USE [Proyeksi_Dana]
GO
/****** Object:  StoredProcedure [dbo].[SUMReportHistoryCashflow]    Script Date: 21/07/2023 13:49:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROCEDURE [dbo].[SUMReportHistoryCashflow]
    @kategori varchar(max),
    @subkategori varchar(max),
	@type varchar(10)
AS
BEGIN
    WITH TanggalProyeksi AS (
        SELECT DATEADD(DAY, number, CONVERT(date, GETDATE())) AS Tanggal
        FROM master.dbo.spt_values
        WHERE type = 'P'
            AND DATEADD(DAY, number, CONVERT(date, GETDATE())) <= DATEADD(YEAR, 1, CONVERT(date, GETDATE()))
            AND DATEPART(WEEKDAY, DATEADD(DAY, number, CONVERT(date, GETDATE()))) NOT IN (1, 7) 
    )
    
    SELECT 
        CAST(tp.Tanggal AS date) AS Tanggal, 
        ISNULL(pt.Kategori, '') AS Kategori, 
        ISNULL(pt.SubKategori, '') AS SubKategori,
        ISNULL(REPLACE(SUM(pt.Nominal), '.00', ''),'') AS Nominal,
        ISNULL((
            SELECT REPLACE(CAST(SUM(pt2.Nominal) AS varchar), '.00', '')
            FROM ProTrxFinansial_Log p2
            JOIN ProTrxFinansialItem pt2 ON p2.Data_Id = pt2.ProTrxFinansial_Id
            WHERE p2.TypeTransaksi = @type 
                AND CAST(p2.TanggalProyeksi AS date) = CAST(tp.Tanggal AS date)
        ),0) AS TotalKategori
    FROM TanggalProyeksi tp
    LEFT JOIN ProTrxFinansial_Log p ON CONVERT(date, p.TanggalProyeksi) = CAST(tp.Tanggal AS date)
    LEFT JOIN ProTrxFinansialItem pt ON p.Data_Id = pt.ProTrxFinansial_Id
    WHERE 
        (pt.Kategori IS NULL OR pt.SubKategori IS NULL OR (pt.Kategori = @kategori AND pt.SubKategori = @subkategori))
        AND (p.TypeTransaksi = @type OR p.TypeTransaksi IS NULL)
    GROUP BY CAST(tp.Tanggal AS date), pt.Kategori, pt.SubKategori
    ORDER BY CAST(tp.Tanggal AS date)
END
