<#
.SYNOPSIS
PS script to extract macro to from .xls file

.DESCRIPTION
This script will take xls file as input and extract macro code if any.

.PARAMETER file
Path to xls file

.EXAMPLE
PS > ./Extract-macro.ps1 C:\Sheet1.xls

#>
[CmdletBinding()]
Param (
  [Parameter(Mandatory=$True,Position=0)]
  [string]$file
)
# Heavily edited from https://github.com/enigma0x3/Generate-Macro/blob/master/Generate-Macro.ps1

#Create excel document
$Excel = New-Object -ComObject "Excel.Application"
$ExcelVersion = $Excel.Version

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
