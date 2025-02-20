<#
.SYNOPSIS
    Generates Event Handler Code Snippet for the desired controlType.

.DESCRIPTION
    The Build-HandlerCode function generates an event handler code snippet for the specified control type.
    $Elements is an array of [PSCustomObject] created by the Find-EveryControl function.
    $ControlType is the type of elements in $Elements and is used mostly for naming the handler.

.PARAMETER Elements
    An array of [PSCustomObject] representing the elements created by Find-EveryControl() function.
    [PSCustomObject]@{
        Name = $Element.Name
        Type = $Element.GetType().ToString()
        Text = (Get-ControlContent -Element $Element)
        Element = $Element
    }

.PARAMETER ControlType
    Is the type of elements in $Elements and is used mostly for naming of the handler

.EXAMPLE
    $elements = @()
    $elements += Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.MenuItem'
    $elements += Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.Primitives.ToggleButton'
    Build-HandlerCode -Elements $elements -ControlType System.Windows.Controls.MenuItem

.NOTES
    Author: Brooks Vaughn
    Date: 2025-02-19 18:12:04
#>
Function Build-HandlerCode {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [System.Object[]]$Elements,
        [Parameter(Mandatory = $true)]
        [String]$ControlType
    )

    If ($null -eq $Elements) {
        Write-Warning "Parameter -Elements is empty or missing"
        Return
    }

    If ([String]::IsNullOrEmpty($ControlType)) {
        Write-Host "Parameter -ControlType is empty or missing" -ForegroundColor Red
        Return
    }

    $eol = [System.Environment]::NewLine
    $sb = [System.Text.StringBuilder]::new()
    $handlerName = "`$handler_$($ControlType.Split('.')[-1])_Click"

    # Starting Code
    [void]$sb.AppendLine(@"
$handlerName = {
    Param ([object]`$theSender, [System.EventArgs]`$e)
    Write-Host ("``$handlerName() Item clicked: {0}" -f `$theSender.Name)
    Switch -Regex (`$theSender.Name) {
"@ + $eol)

    # Body Code consists of Regex patterns of the Item Element Name it finds of matching ControlTypes
    ForEach ($element in $Elements) {
        If (-not [String]::IsNullOrEmpty($element.Name)) {
            [void]$sb.AppendLine(@"
            '^$($element.Name)$' {
                Break
            }
"@)
        }
    }

    # Ending Code
    [void]$sb.AppendLine(@"
        default {
            Write-Host ("{0}: {1}({2})" -f `$theSender.Name, `$e.OriginalSource.Name, `$e.OriginalSource.ToString())
        }
    }
}
"@)

    $sb.ToString()
}

<# Usage Example # >
$elements = @()
$elements += Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.MenuItem'
$elements += Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.Primitives.ToggleButton'
Build-HandlerCode -Elements $elements -ControlType 'System.Windows.Controls.MenuItem'
# How to use the same $elements list to Add_Click() Events
$elements.ForEach({$_.Element.Add_Click($handler_MenuItem_Click)})
#>
