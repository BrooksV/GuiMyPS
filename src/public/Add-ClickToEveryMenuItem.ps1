<#
.SYNOPSIS
    Adds a Click event handler to every menu item within a given menu object.

.DESCRIPTION
    The Add-ClickToEveryMenuItem function traverses the visual tree of a given menu object and adds a Click event handler to every menu item found. Specific menu items can be ignored based on their properties.

.PARAMETER MenuObj
    The root menu object to start the search.

.PARAMETER Handler
    The Click event handler to be added to each menu item.

.EXAMPLE
    Add-ClickToEveryMenuItem -MenuObj $menu -Handler $handler

    Traverses the visual tree of the $menu object and adds the $handler Click event handler to every menu item found.

.NOTES
    Author: Brooks Vaughn
    Date: 2025-02-18 14:22:52
#>
Function Add-ClickToEveryMenuItem {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        $MenuObj,
        [Parameter(Mandatory = $true)]
        [ScriptBlock]$Handler
    )

    # Dump Parameter Values
    $PSBoundParameters.GetEnumerator() | ForEach-Object {
        Write-Verbose "$($_.Key) = $($_.Value)"
    }

    If ([String]::IsNullOrEmpty($Handler)) {
        Write-Warning "Parameter -Handler `$Handler cannot be blank"
        Return
    }
    If ([String]::IsNullOrEmpty($MenuObj)) {
        Write-Warning "Parameter -MenuObj `$MenuObj cannot be blank"
        Return
    }

    Write-Verbose ("Menu Object: [{0}] - {1}" -f $MenuObj.Name, $MenuObj.GetType().ToString())
    ForEach ($child in $MenuObj.Items) {
        Write-Verbose ("  {0} - [{1}]" -f $child.Header, $child.GetType().ToString())
        If ($child.HasItems) {
            Add-ClickToEveryMenuItem -MenuObj $child -Handler:$Handler
        } Else {
            If ($child -is 'System.Windows.Controls.MenuItem') {
                If (-not [String]::IsNullOrEmpty($Handler)) {
                    Write-Verbose ("  Adding Click Event For: {0}" -f $child.Content)
                    $child.Add_Click($Handler)
                } Else {
                    $child.Add_Click($null)
                }
            }
        }
    }
}
