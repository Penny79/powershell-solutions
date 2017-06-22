param
(
    [string]$inputPath="$PSScriptRoot\input\",
    [string]$outputPath="$PSScriptRoot\output\"
) 


function Release-Ref ($ref) { 
    
    ([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) -gt 0) | Out-Null
    [System.GC]::Collect() 
    [System.GC]::WaitForPendingFinalizers() 
} 

function MakeCellValuePositive ($worksheet, $rowIndex, $cellIndex) { 
    
    [double]$cellValue = 0.0
    [string]$test = $worksheet.Cells.Item($rowIndex,$cellIndex).Value
    [bool]$isNumeric = [double]::TryParse($worksheet.cells.item($rowIndex,$cellIndex).value2, [ref]$cellValue)
      
    if($isNumeric)
    {
       $worksheet.cells.item($rowIndex,$cellIndex).value2 *= -1
    }
} 

function ProcessFile($file)
{
 try
    {

        Write-Host "Verarbeite Datei:" $file.Name

        # initialisieren der Zugriffsobjekte
        $excel = New-Object -ComObject Excel.Application
        $excel.visible = $false           
        $workbook = $excel.workbooks.open($file.FullName)
        $worksheets = $workbooks.worksheets
        $worksheet = $workbook.worksheets.Item(1)

        # Finde die letzte genutzt Zeile lieber nicht über RowsUsed, weil das anfällig für manuelle Editierungen im Excel ist
        $lastRowUsed = $worksheet.Range("A2500:A3000").find("Kontrollsumme:").Row       

        # Einlesen der Matrix aus Excel als Range, um viele COM interop Zugriffe zu vermeiden
        $range = $worksheet.Cells.Range("D18:F$($lastRowUsed)")
        $array = ($worksheet.Cells.Range("D18:F$($lastRowUsed)").value2)
                
        # Invertiere die Vorzeichen
        for($rowIndex = 1 ; $rowIndex -le $lastRowUsed - 18 ; $rowIndex++)
        {
            $array[$rowIndex,1] *= -1
            $array[$rowIndex,3] *= -1                      
        }
        
        $range.Value2 = $array

        # Korrigiere noch die Kopf-/Summenzellen
        $lastRowofTimeSlices = $lastRowUsed -1

        
        $worksheet.Cells.Item(4,4).Value = "11XENVIAMBILANZD"
        $worksheet.Cells.Item(5,4).Value = "11XVE-TRADING--X"

        $worksheet.Cells.Item(11,4).Formula = "=MAX(D21:D$lastRowofTimeSlices)"
        $worksheet.Cells.Item(11,6).Formula = "=MAX(D21:D$lastRowofTimeSlices)"

        $worksheet.Cells.Item(12,4).Formula = "=SUM(D21:D$lastRowofTimeSlices)/4"        
        $worksheet.Cells.Item(12,6).Formula = "=SUM(D21:D$lastRowofTimeSlices)/4"

        $worksheet.Cells.Item(15,4).Formula = "=SUM(D21:D$lastRowofTimeSlices) /4"
        $worksheet.Cells.Item(15,6).Formula = "=SUM(D21:D$lastRowofTimeSlices) /4"

        $worksheet.Cells.Item($lastRowUsed, 4).Formula = "=SUM(D21:D$lastRowofTimeSlices) /4"
        $worksheet.Cells.Item($lastRowUsed, 6).Formula = "=SUM(D21:D$lastRowofTimeSlices) /4"
        
        $worksheet.Cells.Item(11,2).Formula = "=MAX(C11:F11)"        
        $worksheet.Cells.Item(12,2).Formula = "=MAX(C12:F12)"

        $workbook.Save()
    
      }
      finally
      {  
        $workbook.Close()
        [void]$excel.quit()
        Release-Ref($worksheet)        
        Release-Ref($workbook)
        Release-Ref($excel)
      }    

    Move-Item $file.FullName $outputPath
    Write-Host "Datei verarbeitet"
}

# ------------------------------------------------
#Skript - Main
# ------------------------------------------------

$filesToProcess = Get-ChildItem $inputPath -Filter *.xls
$numberOfnewFiles = $filesToProcess.Count

Write-Host $numberOfnewFiles "neue Datei(en) zur gefunden."

If ($numberOfnewFiles -eq 0) 
{
    Write-Host "Beende Verarbeitung." 
}
else
{
    Write-Host "Starte Verarbeitung"
}
 

$filesToProcess | Foreach-Object {
  ProcessFile($_)
}
    
Write-Host "Verarbeitung abgeschlossen"

