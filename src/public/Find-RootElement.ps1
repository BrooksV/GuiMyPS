<#
.SYNOPSIS
Locates the root element (main form) of a control.

.DESCRIPTION
This function traverses the visual tree to find the root element of a given control.

.PARAMETER Element
The control for which to find the root element.

.EXAMPLE
    $form = Find-RootElement -Element $Button

.NOTES
    Author: Brooks Vaughn
    Date: 2025-02-19 14:29:36
#>

Function Find-RootElement {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [System.Windows.DependencyObject]
        $Element
    )
    While ($Parent = [System.Windows.Media.VisualTreeHelper]::GetParent($Element)) {
        $Element = $Parent
    }
    Return $Element
}
