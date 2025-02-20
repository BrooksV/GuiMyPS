<#
.SYNOPSIS
    Retrieves the content of a given control element.

.DESCRIPTION
    The Get-ControlContent function retrieves the content of a given control element based on its properties, such as Content, Header, Text, or SelectedItem.

.PARAMETER Element
    The control element from which to retrieve the content.

.EXAMPLE
    Get-ControlContent -Element $button

    Retrieves the content of the $button element.

.NOTES
    Author: Brooks Vaughn
    Date: 2025-02-19 18:16:11
#>
Function Get-ControlContent {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        $Element 
    )

    Switch ($Element) {
        { $_.Content }        { Return $_.Content }
        { $_.Header }         { Return $_.Header }
        { $_.Text }           { Return $_.Text }
        { $_.SelectedItem }   { Return $_.SelectedItem }
        Default               { Return $_.Name }
    }
}
