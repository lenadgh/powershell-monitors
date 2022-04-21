$excelfile = "c:\temp\emaillist.xlsx"
$sheetname = "sheet2"
function get-URLlist{
	$Excel = New-Object -ComObject Excel.Application
	## open the spreadsheet
    $source = $excelfile
	$Workbook = $Excel.Workbooks.Open($source,2,$true)
	## select the worksheet in excel file
	$ws = $Workbook.Sheets.Item("$sheetname")
	## create new PSO for email list
	$URLList = @()
	## Calculate number of used range
	$UsedRange = $ws.UsedRange
    ## specify the column
    $URLLIST += $UsedRange.Cells.Item(1,8).EntireColumn.value2 | Select-Object -Skip 1
    return $URLLIST
    ## close workbook
    $workbook.close($false)
    $Excel.Quit()
}

$URLlist = get-URLlist
$HTTPSsites = @()

foreach ($URL in $URLList){
    if($URL.Contains('https://')){
        $HTTPSsites += $URL}
}
$certstatus = @()

foreach ($HTTPSsite in $HTTPSsites) {
[Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
$req = $null
$req = [Net.HttpWebRequest]::Create($httpssite) 
$req.Timeout = '3600'
$req.AllowAutoRedirect = $false
try{
$req.GetResponse() | Out-Null
[datetime]$expiration = [System.DateTime]::Parse($req.ServicePoint.Certificate.GetExpirationDateString())
$output = [PSCustomObject]@{
   'site' = $httpssite
   'Cert End Date' = $expiration
   'Expires in' = ($expiration - $(get-date)).Days
}
$certstatus += $output
}
catch{
"cannot connect to $httpssite"
}
}
$certstatus
