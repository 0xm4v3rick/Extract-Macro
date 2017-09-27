<#
.SYNOPSIS
PS script to extract macro from Excel and Word files

.DESCRIPTION
This script will take Excel/Word file as input and extract macro code if any.
Supported filetypes: xls,xlsm,doc,docm

.PARAMETER file
Path to Excel/Word file

.EXAMPLE
PS > ./Extract-Macro.ps1 C:\Sheet1.xls

#>
[CmdletBinding()]
Param (
  [Parameter(Mandatory=$True,Position=0)]
  [string]$file
)

# Heavily edited from https://github.com/enigma0x3/Generate-Macro/blob/master/Generate-Macro.ps1

function Word {
    #Create Word document
    $Word = New-Object -ComObject "Word.Application"
    $WordVersion = $Word.Version
    $Word.Visible = $False 
    $Word.DisplayAlerts = "wdAlertsNone"

    #Disable Macro Security
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null


    $Document = $Word.Documents.open($file,$true)
    $xlModule = $Document.VBProject.VBComponents

    foreach($module in $xlModule){
        $line = $module.CodeModule.CountOfLines
            if($line -gt 0){
                $code = $module.CodeModule.Lines(1, $line)
                Write-Host "======== Macro Code Start ============"
                $code 
                Write-Host "======== Macro Code End ============"
            }
    }

    #Cleanup
    $Word.Documents.Close()
    $Word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
    $Word = $Null
    #if (ps Word){kill -name Word}

    #Enable Macro Security
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null
}

function Excel{
    #Create excel document
    $Excel = New-Object -ComObject "Excel.Application"
    $ExcelVersion = $Excel.Version
    $Excel.Visible = $False 
    $Excel.DisplayAlerts = "wdAlertsNone"

    #Disable Macro Security
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null


    $Workbook = $Excel.Workbooks.open($file,$true)
    $xlModule = $Workbook.VBProject.VBComponents
    foreach($module in $xlModule){
        $line = $module.CodeModule.CountOfLines
            if($line -gt 0){
                $code = $module.CodeModule.Lines(1, $line)
                Write-Host "======== Macro Code Start ============"
                $code 
                Write-Host "======== Macro Code End ============"
            }
    }

    #Cleanup
    $Excel.Workbooks.Close()
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | out-null
    $Excel = $Null
    if (ps excel){kill -name excel}

    #Enable Macro Security
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null
}

$extn = [IO.Path]::GetExtension($file)
if (($extn -eq ".doc") -or ($extn -eq ".docm"))
{
    Word
}
elseif(($extn -eq ".xls") -or ($extn -eq ".xlsm"))
{
    Excel
}
else {
    Write-Host "Currently cannot check for this filetype..."
}
