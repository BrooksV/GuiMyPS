<#
.SYNOPSIS
Opens a Windows OpenFileDialog to select a file.

.DESCRIPTION
This function opens a Windows OpenFileDialog to allow the user to select a file. 
It supports options for initial directory, file filter, title dialog, and multi-select. 

.PARAMETER InitialDirectory
The initial directory to open the file dialog.

.PARAMETER FileFilter
The file filter to apply in the dialog.

.PARAMETER TitleDialog
The title of the file dialog.

.PARAMETER AllowMultiSelect
Specifies if multiple files can be selected.

.PARAMETER SaveAs
Specifies if the dialog should be a SaveFileDialog.

.EXAMPLE
$fileName = Get-FileFromDialog -fileFilter 'CSV file (*.csv)|*.csv' -titleDialog "Select A CSV File:"

.NOTES
Author: Brooks Vaughn
Date: "22 February 2025"
#>

Function Get-FileFromDialog {
    [CmdletBinding()] 
    Param (
        [Parameter(Position=0)]
        [string]$InitialDirectory = './',
        
        [Parameter(Position=0)]
        [string]$InitialFileName,
        
        [Parameter(Position=1)]
        [string]$FileFilter = 'All files (*.*)| *.*',
        
        [Parameter(Position=2)] 
        [string]$TitleDialog = '',
        
        [Parameter(Position=3)] 
        [switch]$AllowMultiSelect = $false,
        
        [switch]$SaveAs
    )

    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

    If (-Not $SaveAs) {
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    } Else {
        $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        # $OpenFileDialog.ValidateNames = $false
        $OpenFileDialog.OverwritePrompt = $false
    }

    $OpenFileDialog.InitialDirectory = $InitialDirectory
    $OpenFileDialog.FileName = $InitialFileName
    $OpenFileDialog.Filter = $FileFilter
    $OpenFileDialog.Title = $TitleDialog
    $OpenFileDialog.ShowHelp = If ($Host.Name -eq 'ConsoleHost') { $true } Else { $false }
    
    If ($AllowMultiSelect) { $OpenFileDialog.MultiSelect = $true } 
    
    $OpenFileDialog.ShowDialog() | Out-Null
    
    If ($AllowMultiSelect) { 
        Return $OpenFileDialog.Filenames 
    } Else { 
        Return $OpenFileDialog.Filename 
    }
}
