<#
.SYNOPSIS
PS script to extract macro to from Excel and Word files. Also checks the macro for suspecious/malicious code patterns

.DESCRIPTION
This script will take Excel/Word file as input and extract macro code if any.
Supported filetypes: xls,xlsm,doc,docm

.PARAMETER file
Path to Excel/Word file

.NOTES
  Version:        0.3
  Author:         0xm4v3rick (Samir Gadgil)

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
    
    Try{
        
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
                    Write-Host "======== Macro Code Start ============" -foregroundcolor "green"
                    $code 
                    Write-Host "======== Macro Code End ============" -foregroundcolor "green"
    
                    # Detecting malicous code in the Macro
                    Detection($code)

                }
        }

    
        #Cleanup
        $Word.Documents.Close()
        $Word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
        $Word = $Null
        #if (ps Word){kill -name Word}
    }
    Catch
    {
        $ErrorMessage = $_.Exception.Message
        $ErrorMessage
        $FailedItem = $_.Exception.ItemName
        $FailedItem
    }

    #Enable Macro Security
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null
}

function Excel{

    Try{    

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
                    Write-Host "======== Macro Code Start ============" -foregroundcolor "green"
                    $code 
                    Write-Host "======== Macro Code End ============" -foregroundcolor "green"

                    # Detecting malicous code in the Macro
                    Detection($code)

                }
        }
    
        #Cleanup
        $Excel.Workbooks.Close()
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | out-null
        $Excel = $Null
        if (ps excel){kill -name excel}
    }

    Catch
    {
        $ErrorMessage = $_.Exception.Message
        $ErrorMessage
        $FailedItem = $_.Exception.ItemName
        $FailedItem
    }

    #Enable Macro Security
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null
}

function Detection($vba){

    $keywords = @{"chr\(" = "Use of Char encoding";"Shell"="Use of shell function";"schtasks"="scheduled tasks invocation. Possible backdoor";"Document_Open"="Auto run macro Document_Open";"Auto_Open"="Auto run macro Auto_Open";"(?:[A-Za-z0-9+/]{4}){1,}(?:[A-Za-z0-9+/]{2}[AEIMQUYcgkosw048]=|[A-Za-z0-9+/][AQgw]==)?"="base64 encoded strings [false positive prone]";"(?:[A-Za-z0-9+/]{4}){1,}(?:[A-Za-z0-9+/]{2}[AEIMQUYcgkosw048]=|[A-Za-z0-9+/][AQgw]==)"="base64 encoded strings [Confirmed]"}
    
    $tabName = "SampleTable"

    #Create Table object
    $table = New-Object system.Data.DataTable “$tabName”

    #Define Columns
    $col1 = New-Object system.Data.DataColumn Checks_for,([string])
    $col2 = New-Object system.Data.DataColumn Count,([string])
    #$col3 = New-Object system.Data.DataColumn Instances,([string])

    #Add the Columns
    $table.columns.add($col1)
    $table.columns.add($col2)
    #$table.columns.add($col3)

    foreach($keyword in $keywords.Keys){
        
        $value = $keywords[$keyword]
        $Matches = Select-String -InputObject $vba -Pattern $keyword -AllMatches
        #Write-Host "========  $keyword count ============"
        #$Matches.Matches.Count
        #$Matches

        #Create a row
        $row = $table.NewRow()

        #Enter data in the row
        $row.Checks_for = $value
        $row.Count =   $Matches.Matches.Count 
        #$row.Instances =   $Matches 

        #Add the row to the table
        $table.Rows.Add($row)

    }
    
    Write-Host "======== Suspecious Macro Code Patterns ============" -foregroundcolor "green"

    $table | format-table -Wrap #-AutoSize  

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
    Write-Host "Currently cannot check for this filetype..." -foregroundcolor "red"
    exit
}

