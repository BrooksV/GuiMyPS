<#
.SYNOPSIS
    Adds event handlers to every checkbox within a given element.

.DESCRIPTION
    The Add-EventsToEveryCheckBox function traverses the visual tree of a given element and adds event handlers to every checkbox found. Specific event handlers can be added based on the properties of the checkboxes.

.PARAMETER Element
    The root element to start the search.

.PARAMETER ClickHandler
    The Click event handler to be added to each checkbox.

.PARAMETER CheckedHandler
    The Checked event handler to be added to each checkbox.

.PARAMETER UncheckedHandler
    The Unchecked event handler to be added to each checkbox.

.PARAMETER PreviewMouseUpHandler
    The PreviewMouseUp event handler to be added to each TreeViewItem.

.PARAMETER PreventSelectionScrolling
    A switch to prevent horizontal content scrolling when an item is clicked.

.EXAMPLE
    Add-EventsToEveryCheckBox -Element $window -ClickHandler $clickHandler

    Traverses the visual tree of the $window element and adds the $clickHandler Click event handler to every checkbox found.

.NOTES
    Author: Brooks Vaughn
    Date: 2025-02-18 14:22:52
#>
Function Add-EventsToEveryCheckBox {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        $Element,
        [Parameter(Mandatory = $false)]
        [System.Windows.RoutedEventHandler]$ClickHandler = $null,
        [Parameter(Mandatory = $false)]
        [System.Windows.RoutedEventHandler]$CheckedHandler = $null,
        [Parameter(Mandatory = $false)]
        [System.Windows.RoutedEventHandler]$UncheckedHandler = $null,
        [Parameter(Mandatory = $false)]
        [System.Windows.Input.MouseButtonEventHandler]$PreviewMouseUpHandler = $null,
        [Switch]$PreventSelectionScrolling
    )

    # Dump Parameter Values
    # $PSBoundParameters.GetEnumerator() | ForEach-Object {
    #     Write-Verbose "$($_.Key) = $($_.Value)"
    # }

    Write-Verbose ("Add-EventsToEveryCheckBox(): [{0}] : ({1})" -f $($Element.GetType().ToString()), $($Element.Name))

    If ($Element -is 'System.Windows.Controls.TreeViewItem') {
        $Element.IsExpanded = $true # "True"
        If ($null -ne $PreviewMouseUpHandler) {
            $Element.Add_PreviewMouseUp($PreviewMouseUpHandler)
        }
        If ($PreventSelectionScrolling) {
            # This prevents horizontal content scrolling when an item is clicked
            $Element.Add_RequestBringIntoView({
                Param(
                    [object]$theSender, 
                    [System.Windows.RequestBringIntoViewEventArgs]$e
                )
                Write-Verbose ("`$Element.Add_RequestBringIntoView {0}: {1}({2})" -f $theSender.Name, $e.Source.Name, $e.ToString())
                # Mark the event as handled
                $e.Handled = $true
            })
        }
        <#
        $Count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Element)
        For ($i=0; $i -lt $Count; $I++) {
            $current = [System.Windows.Media.VisualTreeHelper]::GetChild($Element, $i);
            Add-EventsToEveryCheckBox -Element $current -ClickHandler:$ClickHandler -CheckedHandler:$CheckedHandler -UncheckedHandler:$UncheckedHandler
        }
        #>
        Write-Verbose ("  Header Name: {0}, Type: {1}" -f $Element.Header.Type, $Element.Header.Name)
        <#
        If ($Element.Header.Type -eq 'Server') {
            $current = [System.Windows.Media.VisualTreeHelper]::GetChild($Element, 0);
            Add-EventsToEveryCheckBox -Element $current -ClickHandler:$ClickHandler -CheckedHandler:$CheckedHandler -UncheckedHandler:$UncheckedHandler -PreviewMouseUpHandler:$PreviewMouseUpHandler
        }
        #>
    } ElseIf ($Element -is 'System.Windows.Controls.TextBlock') {
        Write-Verbose ("  TextBlock: {0}" -f $($Element.Inlines.Text | Out-String))
    } ElseIf ($Element -is 'System.Windows.Controls.CheckBox') {
        Write-Verbose ("  Adding CheckBox Event Handlers For: " + $Element.Content.Inlines.text -join "")
        If (-not [Sting]::IsNullOrEmpty($ClickHandler)) {
            $Element.Add_Click($ClickHandler)
        }
        If (-not [Sting]::IsNullOrEmpty($CheckedHandler)) {
            $Element.Add_Checked($CheckedHandler)
        }
        If (-not [Sting]::IsNullOrEmpty($UncheckedHandler)) {
            $Element.Add_Unchecked($UncheckedHandler)
        }
    }

    $Count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Element)
    For ($i=0; $i -lt $Count; $i++) {
        $current = [System.Windows.Media.VisualTreeHelper]::GetChild($Element, $i)
        Add-EventsToEveryCheckBox -Element $current -ClickHandler:$ClickHandler -CheckedHandler:$CheckedHandler -UncheckedHandler:$UncheckedHandler -PreviewMouseUpHandler:$PreviewMouseUpHandler
    }
}
