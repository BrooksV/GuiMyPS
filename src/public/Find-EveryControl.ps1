<#
.SYNOPSIS
    Finds every control of a specified type within a given element.

.DESCRIPTION
    The Find-EveryControl function traverses the visual or logical tree of a given element and finds every control of a specified type. Specific controls can be excluded based on their properties.

.PARAMETER ControlType
    The type of control to Gather

.PARAMETER Element
    The root element to start searching from

.PARAMETER ExcludeElement
    A switch to exclude attaching the element to the results.

.PARAMETER UseVisualTreeHelper
    A switch to use VisualTreeHelper for the search.

.PARAMETER IncludeAll
    A switch to include all controls in the search results.

.EXAMPLE
    Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.Primitives.ToggleButton'

    Finds every ToggleButton control within the $form element.

.EXAMPLE
    Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.MenuItem' -ExcludeElement -UseVisualTreeHelper -IncludeAll
    
    Finds all controls (-IncludeAll) using VisualTreeHelper. Best use after exiting ShowDialog(). Excludes attaching Element (-ExcludeElement) to result
    
.EXAMPLE
    $elements = @()
    $elements += Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.Primitives.ToggleButton'
    $elements += Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.MenuItem'
    $elements.ForEach({$_.Element.Add_Click($handler_MenuItem_Click)})

    Find all MenuItem and ToggleButton elements and Add Click Event Handler

.EXAMPLE
    $elements = @()
    $elements += Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.Primitives.ToggleButton' -ExcludeElement
    $elements += Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.MenuItem' -ExcludeElement
    Build-HandlerCode -Elements $elements -ControlType System.Windows.Controls.MenuItem
    
    Find all MenuItem and ToggleButton elements and Generate the code for the Click Event Handler

.NOTES
    Author: Brooks Vaughn
    Date: 2025-02-19 14:29:36
#>
Function Find-EveryControl {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        $ControlType,
        [Parameter(Mandatory = $true)]
        $Element,
        [Parameter(Mandatory = $false)]
        [Switch]$ExcludeElement,
        [Parameter(Mandatory = $false)]
        [Switch]$UseVisualTreeHelper,
        [Parameter(Mandatory = $false)]
        [Switch]$IncludeAll
    )

    If (-not [String]::IsNullOrEmpty($IncludeAll) -and $IncludeAll) {
        $includeElement = $null
        If (-not $ExcludeElement -or [String]::IsNullOrEmpty($ExcludeElement)) {
            $includeElement = $Element
        }
        [PSCustomObject]@{
            Name = $Element.Name
            Type = $Element.GetType().ToString()
            Text = (Get-ControlContent -Element $Element)
            Element = $includeElement
        }        
    } Else {
        If ($Element -is $ControlType) {
            $includeElement = $null
            If (-not $ExcludeElement -or [String]::IsNullOrEmpty($ExcludeElement)) {
                $includeElement = $Element
            }
            [PSCustomObject]@{
                Name = $Element.Name
                Type = $Element.GetType().ToString()
                Text = (Get-ControlContent -Element $Element)
                Element = $includeElement
            }
        }
    }

    If ($UseVisualTreeHelper) {
        $Count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Element)
        If ($Count -gt 0) {
            For ($i=0; $i -lt $Count; $i++) {
                $current = [System.Windows.Media.VisualTreeHelper]::GetChild($Element, $i)
                Find-EveryControl -Element $current -ControlType:$ControlType -IncludeAll:$IncludeAll -ExcludeElement:$ExcludeElement -UseVisualTreeHelper:$UseVisualTreeHelper
            }
        }
    } Else {
        If ($Element.HasContent) {
            ForEach ($item in $Element.Content) {
                Find-EveryControl -Element $item -ControlType:$ControlType -IncludeAll:$IncludeAll -ExcludeElement:$ExcludeElement
            }
        }
        If ($Element.Children) {
            ForEach ($child in $Element.Children) {
                Find-EveryControl -Element $child -ControlType:$ControlType -IncludeAll:$IncludeAll -ExcludeElement:$ExcludeElement
            }
        }
        If ($Element.HasItems) {
            ForEach ($item in $Element.Items) {
                Find-EveryControl -Element $item -ControlType:$ControlType -IncludeAll:$IncludeAll -ExcludeElement:$ExcludeElement
            }
        }
    }
}

# Example usage:
# Find-EveryControl -Element $WPF_menuMasterDataGrid -ControlType 'System.Windows.Controls.MenuItem'
# Find-EveryControl -Element $WPF_menuDetailDataGrid -ControlType 'System.Windows.Controls.MenuItem'
# Find-EveryControl -Element $WPF_tabControl -ControlType 'System.Windows.Controls.Button'
# Find-EveryControl -Element $WPF_gridSqlQueryEditor -ControlType 'System.Windows.Controls.Primitives.ToggleButton'
# Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.MenuItem'
# Find-EveryControl -Element $WPF_tabControl -ControlType 'System.Windows.Controls.Primitives.ToggleButton'
# Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.Primitives.ToggleButton'
# Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.MenuItem'
# Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.RadioButton'
# Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.Combobox'
# Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.Button'
# Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.Primitives.ToggleButton'
# Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.MenuItem' -IncludeAll -ExcludeElement -UseVisualTreeHelper
# Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.MenuItem' -IncludeAll -ExcludeElement

# How to identify all controls (-IncludeAll) from VisualTreeHelper. Use after exiting ShowDialog()
# Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.MenuItem' -ExcludeElement -UseVisualTreeHelper -IncludeAll
# Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.MenuItem' -ExcludeElement -UseVisualTreeHelper
# Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.MenuItem' -ExcludeElement

