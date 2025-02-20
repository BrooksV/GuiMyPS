<#
.SYNOPSIS
    Returns a PSObject containing details about an object's properties.

.DESCRIPTION
    The Get-ObjectPropertyDetail function returns a PSObject that lists the properties of the input object along with their names, values, and types.

.PARAMETER InputObject
    The object whose properties are to be detailed.

.EXAMPLE
    $details = Get-ObjectPropertyDetail -InputObject $form1
    $details | Format-Table -AutoSize
    $details | Format-Table -AutoSize -Force -Wrap -Property Property, Value, Type

    Retrieves the properties of the object $form1 and displays them in a formatted table.

.NOTES
    Author: Brooks Vaughn
    Date: 2025-02-18 15:41:05
#>
Function Get-ObjectPropertyDetail {
    [CmdletBinding()]
    Param (
        [parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [Object]$InputObject
    )
    Begin {
        $ReturnObject = [System.Collections.Generic.List[PSCustomObject]]::new()
    }
    Process {
        ForEach ($property in $InputObject.PSObject.Properties) {
            $typeName = [Microsoft.PowerShell.ToStringCodeMethods]::Type([type]$property.TypeNameOfValue)
            $value = $property.Value

            If ($typeName -like '*IScriptExtent*') {
                $file = if ($null -eq $value.File) { "" } else { Split-Path -Leaf $value.File }
                $value = "{0} ({1},{2})-({3},{4})" -f $file, $value.StartLineNumber, $value.StartColumnNumber, $value.EndLineNumber, $value.EndColumnNumber
            }

            [void]$ReturnObject.Add([PSCustomObject]@{
                Property = $property.Name
                Value    = $value
                Type     = $typeName
            })
        }
    }
    End {
        $ReturnObject
    }
}

