<#
.SYNOPSIS
    Adds a Click event handler to every button within a given element.

.DESCRIPTION
    The Add-ClickToEveryButton function traverses the visual tree of a given element and adds a Click event handler to every button found. Specific buttons can be ignored based on their names.

.PARAMETER Element
    The root element to start the search.

.PARAMETER ClickHandler
    The Click event handler to be added to each button.

.EXAMPLE
    Add-ClickToEveryButton -Element $window -ClickHandler $ClickHandler

    Traverses the visual tree of the $window element and adds the $ClickHandler Click event handler to every button found.

.NOTES
    Author: Brooks Vaughn
    Date: 2025-02-18 14:22:52
#>
Function Add-ClickToEveryButton {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        $Element,
        [Parameter(Mandatory = $false)]
        [System.Windows.RoutedEventHandler]$ClickHandler
    )

    # Dump Parameter Values
    # $PSBoundParameters.GetEnumerator() | ForEach-Object {
    #     Write-Verbose "$($_.Key) = $($_.Value)"
    # }

    If ($Element -is 'System.Windows.Controls.Button') {
        Write-Verbose ("Add-ClickToEveryButton() Button Object: [{0}] - {1}" -f $Element.GetType().ToString(), $Element.Name)
        If (-not [String]::IsNullOrEmpty($ClickHandler)) {
            Write-Verbose ("  Adding Click Event For: {0}" -f $Element.Content)
            $Element.Add_Click($ClickHandler)
        }
    } Else {
        Write-Verbose ("Add-ClickToEveryButton() Object: [{0}] - {1}" -f $Element.GetType().ToString(), $Element.Name)
    }

    If ($Element.HasItems -or $Element.HasContent -or $Element.Child.Count -gt 0 -or $Element.Children.Count -gt 0 -or $Element.Items.Count -gt 0) {
        # The logical tree can contain any type of object, not just 
        # instances of DependencyObject subclasses. LogicalTreeHelper
        # only works with DependencyObject subclasses, so we must be
        # sure that we do not pass it an object of the wrong type.
        $depObj = $Element
        If ($null -ne $depObj) {
            ForEach ($logicalChild in ([System.Windows.LogicalTreeHelper]::GetChildren($depObj))) {
                Add-ClickToEveryButton -Element $logicalChild -ClickHandler:$ClickHandler
            }
        }
    }
}

