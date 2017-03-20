
$d = get-date -format yyyyMMdd
$infn = $d + "_OOSExtract.csv"
$outfn = $d + "_OOSExtract.zip"
$daysback = 8


$connString = "Server=oosdbprd;Database=oos;Trusted_Connection=True;" 


$query = @"
        SELECT  r.REGION_ABBR AS Region ,
        s.STORE_ABBREVIATION AS StoreAbbr ,
        s.PS_BU ,
        OffsetCorrectedCreateDate AS OosScannedOn ,
        brand_name AS Brand ,
        UPC ,
        LONG_DESCRIPTION AS ItemDescription ,
        PS_TEAM AS Team ,
        PS_SUBTEAM AS Subteam ,
		Category_name AS Category, 
		Class_name AS Class,
        CAST(VENDOR_KEY AS VARCHAR(100)) AS Vendor_Key
FROM    region r
        INNER JOIN store s ON r.ID = s.REGION_ID
        INNER JOIN dbo.REPORT_HEADER rh ON s.ID = rh.STORE_ID
        INNER JOIN dbo.REPORT_DETAIL rd ON rh.ID = rd.REPORT_HEADER_ID
WHERE   OffsetCorrectedCreateDate > DATEADD(Day, -$daysback, GETDATE())
ORDER BY STORE_ABBREVIATION ,
        OffsetCorrectedCreateDate DESC
"@;

#$query = "SELECT  r.REGION_ABBR AS Region,        s.STORE_ABBREVIATION AS StoreAbbr,        s.PS_BU ,        OffsetCorrectedCreateDate AS OosScannedOn ,        brand_name AS Brand,        UPC ,        LONG_DESCRIPTION AS ItemDescription,        PS_TEAM AS Team,        PS_SUBTEAM AS Subteam,		CAST(VENDOR_KEY AS VARCHAR(100))AS Vendor_Key FROM    region r        INNER JOIN store s ON r.ID = s.REGION_ID        INNER JOIN dbo.REPORT_HEADER rh ON s.ID = rh.STORE_ID        INNER JOIN dbo.REPORT_DETAIL rd ON rh.ID = rd.REPORT_HEADER_ID WHERE   OffsetCorrectedCreateDate >= '1/23/2017' and OffsetCorrectedCreateDate < '1/30/2017' ORDER BY STORE_ABBREVIATION, OffsetCorrectedCreateDate desc"
#$query = "SELECT  r.REGION_ABBR AS Region ,        s.STORE_ABBREVIATION AS StoreAbbr ,        s.PS_BU ,        OffsetCorrectedCreateDate AS OosScannedOn ,        brand_name AS Brand ,        UPC ,        LONG_DESCRIPTION AS ItemDescription ,        PS_TEAM AS Team ,        PS_SUBTEAM AS Subteam ,        CAST(VENDOR_KEY AS VARCHAR(100)) AS Vendor_Key FROM    region r        INNER JOIN store s ON r.ID = s.REGION_ID        INNER JOIN dbo.REPORT_HEADER rh ON s.ID = rh.STORE_ID        INNER JOIN dbo.REPORT_DETAIL rd ON rh.ID = rd.REPORT_HEADER_ID WHERE   OffsetCorrectedCreateDate >= '8/1/2016' AND OffsetCorrectedCreateDate < '8/8/2016'        ORDER BY STORE_ABBREVIATION ,        OffsetCorrectedCreateDate DESC"
#$query = "SELECT  r.REGION_ABBR AS Region,        s.STORE_ABBREVIATION AS StoreAbbr,        s.PS_BU ,        OffsetCorrectedCreateDate AS OosScannedOn ,        brand_name AS Brand,        UPC ,        LONG_DESCRIPTION AS ItemDescription,        PS_TEAM AS Team,        PS_SUBTEAM AS Subteam,		CAST(VENDOR_KEY AS VARCHAR(100))AS Vendor_Key FROM    region r        INNER JOIN store s ON r.ID = s.REGION_ID        INNER JOIN dbo.REPORT_HEADER rh ON s.ID = rh.STORE_ID        INNER JOIN dbo.REPORT_DETAIL rd ON rh.ID = rd.REPORT_HEADER_ID WHERE   OffsetCorrectedCreateDate > DATEADD(Day,-$daysback,getdate()) ORDER BY STORE_ABBREVIATION, OffsetCorrectedCreateDate desc"


$cn = New-Object System.Data.SqlClient.SqlConnection($connString)
$cn.open()            
$readcmd = New-Object system.Data.SqlClient.SqlCommand
$readcmd.Connection = $cn
$readcmd.CommandTimeout = '300'
$readcmd.CommandText = $query
$da = New-Object system.Data.SqlClient.SqlDataAdapter($readcmd)            
$dt = New-Object system.Data.DataTable            
[void]$da.fill($dt)            
$cn.Close()

$dt | export-csv $infn -notypeinformation
write-zip $infn $outfn
Remove-Item $infn
