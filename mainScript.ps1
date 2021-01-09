
# checks if outlook or excel processes are active

$OutlookP = Get-Process Outlook -ErrorAction SilentlyContinue
$ExcelP = Get-Process Excel -ErrorAction SilentlyContinue

Write-Host "Checking if outlook or excel are running"

if ($OutlookP) {
  # try gracefully first
  $OutlookP.CloseMainWindow()
  # kill after five seconds
  Sleep 5
  if (!$OutlookP.HasExited) {
    $OutlookP | Stop-Process -Force
  }
  Remove-Variable OutlookP
}

if ($ExcelP) {
  # try gracefully first
  $ExcelP.CloseMainWindow()
  # kill after five seconds
  Sleep 5
  if (!$ExcelP.HasExited) {
    $ExcelP | Stop-Process -Force
  }
  Remove-Variable ExcelP
}

Write-Host "Starting Excel process"

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true

#avoid saving popups
$Excel.DisplayAlerts = $false

Write-Host "Setting up variables"
###VARIABLES###

#dates
$fecha= Get-Date -format "dd\\ MM\\ yyyy"
$anio= Get-Date -format "yyyy"
$mesanio= date -Format y
$mes= Get-Date -Uformat %B
$mesnum= Get-Date -Uformat %m

#MODIFY THESE VARIABLES TO SUITE YOUR NEEDS
#Rutas
$rutacarpeta="route to pdfs folder"
$Path = "route to excel file"

$emaillocal1= "destination email"
$namelocal1="testname"
$cadenaCheckslocal1="testcheck"

Write-Host "Opening excel"
#open excel de path a wb
$wb = $Excel.Workbooks.Open($Path)

$xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type]
$xlQuality = "Microsoft.Office.Interop.Excel.xlQualityStandard" -as [type]

Write-Host "Modifing EXCEL"

#THE MODIFICATIONS ARE LOCATED WITH XY CELL COORDINATES
#MODIFY TO SUITE YOUR NEEDS

$wb.Worksheets.Item(1).Copy($wb.Worksheets.Item(1))
$newSheet = $wb.Worksheets.Item(1)
$newSheet.Activate()
$lastSheet = $wb.WorkSheets.Item($wb.WorkSheets.Count) 
$newSheet.Move([System.Reflection.Missing]::Value, $lastSheet)
$newSheet.Cells.Item(15,1) =$fecha
$cadena="Periodo: $mesanio"
$newSheet.Cells.Item(20,2) =$cadena
$num= 2*$mesnum - 1
$newSheet.Cells.Item(11,1)=$num
$newSheet.Cells.Item(11,2)=$anio
$namesheet = "$namelocal1$num-$anio"
$newSheet.Name = $namesheet

Write-Host "Exporting to pdf"

$fechapdf= Get-Date -format MM-yyyy
$xlFromPage= $wb.Worksheets.Count
$xlToPage= $wb.Worksheets.Count
$wb.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, "$rutacarpeta\$namelocal1 $fechapdf.pdf", $xlQuality, $false, $true, $xlFromPage, $xlToPage)
$wb.SaveAs($Path)
$wb.Close($false)

$Excel.Quit()

Sleep 2

Write-Host "Closing excel and starting outlook with specific profile"

Start-Process Outlook -ErrorAction SilentlyContinue -ArgumentList '/profile "Outlook" '
$Outlook = New-Object -ComObject Outlook.Application

Sleep 2

Write-Host "Crafting emails"

#email a local1
$Mail = $Outlook.CreateItem(0)
$Mail.Sender ="SENDER EMAIL"
$Mail.To = $emaillocal1
$Mail.Subject = "TEST SUBJECT"
$Mail.Body ="TEST BODY."
$Mail.Attachments.Add("$rutacarpeta\$namelocal1 $fechapdf.pdf")
Sleep 1

Write-Host "Checking pdfs integrity before sending"

#variables de iTextSharp
Add-Type -Path "YOUR ROUTE\itextsharp.5.5.13\lib\itextsharp.dll"
$pdfs = @("$rutacarpeta\$namelocal1 $fechapdf.pdf")
$results = @()
$keywords = @("$cadenaCheckslocal1")


Write-Host "processing -" $pdfs[0]`n-------------------------------

$reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $pdfs[0]

for($page = 1; $page -le $reader.NumberOfPages; $page++) {
     # set the page text
    $pageText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader,$page).Split([char]0x000A)
    Write-Host "Matching "$keywords[0]" with " $pageText
}
if($pageText -match $keywords[0]) {
    $results += 1
}
$reader.Close()

Write-Host "Results of integrity checks: "$results

if ($results[0] -eq 1 ){
    Write-Host "EVERYTHING SEEMS OK, SENDING EMAIL"
    $Mail.Send()
    #PRINTING PDFS
    Write-Host "Printing pdfs...."
    #start-process -filepath "$rutacarpeta\$namelocal1 $fechapdf.pdf" -verb print
}else{
    Write-Host "something is wrong, will send email to administrator"
    $Mail4 = $Outlook.CreateItem(0)
    $Mail4.Sender ="test email"
    $Mail4.To = "test administrator"
    $Mail4.Subject = "error"
    $Mail4.Body ="Se ha producido un error, salida de results: $results"
    Sleep 1
    $Mail4.Send()
}


Write-Host "waiting 20 secs to quit outlook until mails are sended"
Sleep 20
$Outlook.Quit()




