<#
.SYNOPSIS
Saves the content of a DataGrid to various formats such as Clipboard, CSV, or Excel.

.DESCRIPTION
This function converts the content of a DataGrid and saves it to the specified format.
It supports saving to Clipboard, CSV files, and Excel files.

.PARAMETER InputObject
The object to convert and save. This can be a DataGrid or DataRowView.

.PARAMETER Path
The path and filename of the SaveAs file.

.PARAMETER SaveAs
The target output format. This can be Clipboard, CSV, or Excel.

.EXAMPLE
Save-DataGridContent -InputObject $WPF_dataGridSqlQuery -SaveAs CVS -Verbose
Save-DataGridContent -InputObject $WPF_dataGridSqlQuery -SaveAS CSV -Path C:\Temp\dataGridSqlQuery.csv
$WPF_dataGridSqlQuery | Save-DataGridContent -SaveAs Excel -Verbose
Save-DataGridContent -InputObject $WPF_dataGridSqlQuery -SaveAs Clipboard -Verbose

.NOTES
Author: Brooks Vaughn
Date: "22 February 2025"
#>

Function Save-DataGridContent {
    [CmdletBinding()] 
    Param (
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, HelpMessage='Object to convert and save')]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSObject]$InputObject,

        [Parameter(Mandatory=$false, Position=1, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, HelpMessage='Path and Filename of SaveAs File')]
        [ValidateScript({
            $Parent = Split-Path $_ -Parent -ErrorAction SilentlyContinue
            If (-Not (Test-Path -Path $Parent -PathType Container -ErrorAction SilentlyContinue)) {
                Throw "Specify a valid path. Parent '$Parent' does not exist: $_"
            }
            $true
        })]        
        [string]$Path,

        [Parameter(Position=2, Mandatory=$false, HelpMessage='Target output format')]
        [ValidateSet('Clipboard', 'CSV', 'Excel')]
        [string]$SaveAs = 'Clipboard'
    )
    Begin {
        Add-Type -AssemblyName System.Windows.Forms
        $RowCount = 0
        # $PipelineInput = -Not $PSBoundParameters.ContainsKey('InputObject')
        # $Index = 1
        $OutObject = @()
    }
    Process {
        If ($SaveAs -eq 'clipboard' -and $InputObject -is 'System.Windows.Controls.DataGrid') {
            ForEach ($Row in $InputObject.Items) {
                If ($Row.ToString() -eq '{NewItemPlaceholder}') {
                    Continue
                }
                $RowCount++
                If ($RowCount -ge $InputObject.Items.Count) {
                    $Row
                    Break
                }
                $OutObject += "Record # {0}{1}" -f $RowCount, [System.Environment]::NewLine
                ForEach ($Column in $InputObject.Columns.Header) {
                    $OutObject += "{0} : {1}{2}" -f $Column, ($Row."$Column"), [System.Environment]::NewLine
                }
                $OutObject += [System.Environment]::NewLine
            }
            $OutObject += "Total Record Count: {0}{1}" -f $RowCount, [System.Environment]::NewLine
        } ElseIf ($SaveAs -eq 'clipboard' -and $InputObject -is 'System.Data.DataRowView') {
            ForEach ($Row in $InputObject) {
                $RowCount++
                $OutObject += "Record # {0}{1}" -f $RowCount, [System.Environment]::NewLine
                ForEach ($Column in (($Row."Row").PSObject.Properties.Name | Where-Object {$_ -notin @('RowError','RowState','Table','ItemArray','HasErrors')})) {
                    $OutObject += "{0} : {1}{2}" -f $Column, ($Row."$Column"), [System.Environment]::NewLine
                }
                $OutObject += [System.Environment]::NewLine
            }
            $OutObject += "Total Record Count: {0}{1}" -f $RowCount, [System.Environment]::NewLine
        }
    }
    End {
        Switch ($SaveAs) {
            'clipboard' {
                [System.Windows.Forms.Clipboard]::SetText($OutObject)
                Write-Host ("Record Copied to {0}" -f $SaveAs)
            }
            'CSV' {
                $SaveFile = Get-FileFromDialog -InitialDirectory $Parent -InitialFileName $Path -TitleDialog "Specify SaveAs $SaveAs path and filename" -FileFilter 'CSV files (*.csv)|*.csv|CSV files (*.txt)|*.txt|All files (*.*)|*.*' -SaveAs
                If (-Not [string]::IsNullOrEmpty($SaveFile)) {
                    $SaveFile = Get-UniqueFileName -Name $SaveFile
                    $InputObject.Items.Where({$_.ToString() -ne '{NewItemPlaceholder}'}) | Export-Csv -Path $SaveFile -Force -NoTypeInformation
                    Write-Host ("{0} Saved to: {1}" -f $SaveAs, $SaveFile)
                } Else {
                    Write-Host ("SaveAs {0} aborted by user" -f $SaveAs)
                }
            }
            'Excel' {
                $SaveFile = Get-FileFromDialog -InitialDirectory $Path -TitleDialog "Specify SaveAs $SaveAs path and filename" -FileFilter 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*' -SaveAs
                If (-Not [string]::IsNullOrEmpty($SaveFile)) {
                    $XLFile = Get-UniqueFileName -Name $SaveFile
                    Export-XLSX -InputObject $InputObject.Items -Path $XLFile -Table -Autofit -WorksheetName $rsObject.QueryDB
                    Write-Host ("{0} Saved to: {1}" -f $SaveAs, $SaveFile)
                } Else {
                    Write-Host ("SaveAs {0} aborted by user" -f $SaveAs)
                }
            }
            Default {}
        }
    }
}
