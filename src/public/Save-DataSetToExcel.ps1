<#
.SYNOPSIS
Saves a dataset to an Excel file.

.DESCRIPTION
This function converts a dataset to an Excel file and saves it to the specified path.

.PARAMETER InputObject
The dataset to convert and save.

.PARAMETER Path
The path and filename of the SaveAs file.

.EXAMPLE

.NOTES
Author: Brooks Vaughn
Date: "22 February 2025"
#>

Function Save-DatasetToExcel {
    [CmdletBinding()] 
    Param (
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, HelpMessage='Object to convert and save')]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSObject]$InputObject,

        [Parameter(Mandatory=$false, Position=1, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, HelpMessage='Path and Filename of SaveAs File')]
        [ValidateScript({
            $Parent = Split-Path $_ -Parent -ErrorAction SilentlyContinue
            If (-not (Test-Path -Path $Parent -PathType Container -ErrorAction SilentlyContinue)) {
                Throw "Specify a valid path. Parent '$Parent' does not exist: $_"
            }
            $true
        })]        
        [string]$Path
    )

    Begin {
        Initialize-Module -ModuleName "ImportExcel"

        # $RowCount = 0
        # $PipelineInput = -not $PSBoundParameters.ContainsKey('InputObject')
        # $Index = 1
        # $OutObject = @()

        If (-not [string]::IsNullOrEmpty($Path)) {
            # Create unique TimeStamped File pathname for saving output to
            $XLFile = Get-UniqueFileName -Name "$Parent\$Path"
            Write-Host ("Exporting DataSet To: {0}" -f $XLFile)
        } Else {
            Write-Host ("Error: -Path is missing or empty")
            $XLFile = Get-FileFromDialog -initialDirectory $Path -titleDialog "Specify SaveAs $Parent path and filename" -fileFilter 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*' -SaveAs
        }

        # ------------------------------------------------------------------
        # Begin Excel Creation
        # ------------------------------------------------------------------
        Write-Host ("Exporting Query Results to Excel File ($XLFile)")
        $xlApp = New-Object OfficeOpenXml.ExcelPackage $XLFile
        [void]$xlApp.Workbook.Worksheets.Add($WorkSheetName)
        $wbSource = $xlApp.Workbook

        # Update workBook properties 
        $wbSource.Properties.Author = $env:USERNAME
        $wbSource.Properties.Title = "Query Results Export"
        $wbSource.Properties.Subject = "Query Results Export"
        $wbSource.Properties.Company = ""

        # -----------------------------------------------------
        # Create Queries Worksheet and Populate
        # -----------------------------------------------------
        Add-WorkSheet -Workbook $wbSource -Name 'Query' -Headers @{ 'Query Number' = 1; 'SQL Query' = 2 } -HeaderBorderStyle ([OfficeOpenXml.Style.ExcelBorderStyle]::Thin)
        $xlApp.Save()
        Close-ExcelPackage -ExcelPackage $xlApp
        $xlApp.Dispose()
        $xlApp = $null        
    }

    Process {
        # Add each Result Set ($InputObject.Value) as a new worksheet named as ($InputObject.Name)
        Export-XLSX -InputObject $InputObject.Value -Path $XLFile -Table -Autofit -WorksheetName $InputObject.Name
    }

    End {
        Write-Host ("Export Completed")
    }
}
