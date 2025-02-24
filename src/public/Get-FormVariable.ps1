<#
.SYNOPSIS
    Displays WPF_* form variables which are XAML elements whose x:Name begins with an _ underscore character.

.DESCRIPTION
    This function displays a list of WPF_* variables of XAML Elements which can accessed and referenced directly,

.EXAMPLE
    Get-FormVariable
-----------------------------------------------------------------
Found the following intractable elements from our form
-----------------------------------------------------------------
Name                           Value
----                           -----
WPF_dataGrid                   System.Windows.Controls.DataGrid Items.Count:0
WPF_saveButton                 System.Windows.Controls.Button: Save
-----------------------------------------------------------------

.NOTES
    Author: Brooks Vaughn
    Date: Wednesday, 19 February 2025 19:26
#>

Function Get-FormVariable {
    [CmdletBinding()]
    param ()

    if ($null -eq $ReadmeDisplay -or $ReadmeDisplay -eq $false) {
        "If you need to reference this display again, run Get-FormVariables"
    }
    ("`n$("-" * 65)`nGet-FormVariable() Found the following intractable elements from the XAML form`n$("-" * 65)")
    ((Get-Variable WPF*) | Select-Object -Property Name, Value | Out-String).Trim()
    ("$("-" * 65)`n")
}

