<#
.SYNOPSIS
    Helper function that processes and loads XAML (as string, filename, or as [XML] object) into a WPF Form Object.

.DESCRIPTION
    The New-XamlWindow function processes and loads XAML (as string, filename, or as [XML] object) into a WPF Form Object.

.PARAMETER xaml
    The XAML content to process and load.

.PARAMETER NoXRemoval
    A switch to prevent the removal of x:Name attributes.

.EXAMPLE
    # Test for XAML String
    try {
        $form1 = New-XamlWindow -xaml $inputXML
        Add-ClickToEveryButton -Element $form1 -ClickHandler $handler_button_Click
        $form1.ShowDialog()
    } catch {
        Write-Warning ($_ | Format-List | Out-String)
    }

.EXAMPLE
    # Test for File Path to XAML file with two approaches for Adding Click Events
    # Note: Code for $handler_MenuItem_Click can be generated using Build-HandlerCode()
    try {
        $form = New-XamlWindow -xaml 'C:\Git\SqlQueryEditor\src\resources\SqlQueryEditor.xaml'
        # Add-ClickToEveryMenuItem -MenuObj $WPF_menuMasterDataGrid -Handler $handler_MenuItem_Click
        # Add-ClickToEveryMenuItem -MenuObj $WPF_menuDetailDataGrid -Handler $handler_MenuItem_Click
        $elements = @()
        $elements += Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.Primitives.ToggleButton'
        $elements += Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.MenuItem'
        $elements.ForEach({$_.Element.Add_Click($handler_MenuItem_Click)})
        $form.ShowDialog()
    } catch {
        Write-Warning ($_ | Format-List | Out-String)
    }

.EXAMPLE
    # Test for [System.Xml.XmlDocument]
    try {
        $xaml = [xml]$inputXML
        $form2 = New-XamlWindow -xaml $xaml
        Add-ClickToEveryButton -Element $form2 -ClickHandler $handler_button_Click
        $form2.ShowDialog()
    } catch {
        Write-Warning ($_ | Format-List | Out-String)
    }

.NOTES
    Author: Brooks Vaughn
    Date: 2025-02-19 18:28:36
#>
Function New-XamlWindow {
    [CmdletBinding(SupportsShouldProcess = $true)]
    Param (
        [Parameter(Mandatory = $true)]
        [Alias("InputObject")]
        [System.Object]$xaml,

        [Parameter(Mandatory = $false)]
        [switch]$NoXRemoval
    )

    # Load the necessary assemblies
    Add-Type -AssemblyName PresentationFramework

    <#
    XAML Elements might use either x:Name="" or Name="" 
    The x: refers to the namespace xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" which is not automatically defined in PowerShell.
    To use XPath queries during preprocessing of the XAML Text, you have to create a NamespaceManager and add the missing namespaces.
    For x:Name, the XPath query would be $xaml.SelectNodes("//*[@x:Name]", $nsManager).
    For Name, the XPath query would be $xaml.SelectNodes("//*[@Name]", $nsManager).
    By removing x:Name from the XAML string before converting to [System.Xml.XmlDocument], there is no need for the NamespaceManager and XPath needs only $xaml.SelectNodes("//*[@Name]").
    #>
    Function Remove-XName {
        [CmdletBinding(SupportsShouldProcess = $true)]
        Param (
            [Parameter(Mandatory = $true)]
            [string]$xamlString,

            [Parameter(Mandatory = $false)]
            [switch]$NoXRemoval
        )

        If ($NoXRemoval) {
            $xamlString
        } Else {
            If ($PSCmdlet.ShouldProcess("XAML String", "Remove x:Name attributes")) {
                ((($xamlString -replace 'mc:Ignorable="d"', '') -replace "x:Na", 'Na') -replace '^<Win.*', '<Window')
            } Else {
                $xamlString
            }
        }
    }

    # Define return and working variables
    $uiForm = $null
    $xamlReader = $null

    Try {
        # Process the $xaml input object into an XmlDocument and create reader
        If ($xaml -is [System.Xml.XmlDocument]) {
            $newXaml = $xaml
        } ElseIf ($xaml -is [System.String]) {
            If (Test-Path -Path $xaml -PathType Leaf -ErrorAction SilentlyContinue) {
                $xamlString = Remove-XName -NoXRemoval:$NoXRemoval -xamlString (Get-Content -Path $xaml -Raw)
                If ($xamlString -match '<(Window|Grid|StackPanel|Button|DataGrid|TextBox)') {
                    $newXaml = [xml]$xamlString
                } Else {
                    Throw "-xaml parameter as file path was valid but its content is not XAML"
                }
            } ElseIf ($xaml -match '<(Window|Grid|StackPanel|Button|DataGrid|TextBox)') {
                $xamlString = Remove-XName -NoXRemoval:$NoXRemoval -xamlString $xaml
                $newXaml = [xml]$xamlString
            } Else {
                Throw "-xaml parameter as XAML String has no valid XAML content"
            }
        } Else {
            Throw "-xaml is not a valid [System.Xml.XmlDocument], XAML string, or filepath to a XAML file"
        }

        # Perform XML data cleanup and create PowerShell $WPF_* Variables for XAML names that begin with "_"
        # Remove unsupported namespaces created by XAML editors like Blend / Visual Studio
        $newXaml.Window.RemoveAttribute('x:Class')
        $newXaml.Window.RemoveAttribute('d')
        $newXaml.Window.RemoveAttribute('mc:Ignorable')

        $problemObjects = $newXaml.SelectNodes("//*[@Name]").Where({
            $_.Click -or $_.TextChanged -or $_.SelectionChanged -or $_.Checked -or $_.Unchecked -or $_.ValueChanged}) | 
                Select-Object -Property Name, LocalName, Click, TextChanged, SelectionChanged, Checked, Unchecked, ValueChanged

        $errorHeading = "XAML String has unsupported events defined and cannot be converted by PowerShell WPF" + [System.Environment]::NewLine
        $errorText = ForEach ($obj in $problemObjects) {
            Switch ($obj) {
                { $_.Click } { "$($_.LocalName) named ($($_.Name)) has Click event and needs to be removed"; break }
                { $_.TextChanged } { "$($_.LocalName) named ($($_.Name)) has TextChanged event and needs to be removed"; break }
                { $_.SelectionChanged } { "$($_.LocalName) named ($($_.Name)) has SelectionChanged event and needs to be removed"; break }
                { $_.Checked } { "$($_.LocalName) named ($($_.Name)) has Checked event and needs to be removed"; break }
                { $_.Unchecked } { "$($_.LocalName) named ($($_.Name)) has Unchecked event and needs to be removed"; break }
                { $_.ValueChanged } { "$($_.LocalName) named ($($_.Name)) has ValueChanged event and needs to be removed"; break }
                Default { "Skipping: $($_.LocalName) named ($($_.Name))" }
            }
        }
        If ($problemObjects.Count -gt 0) {
            Write-Host ($errorHeading) -ForegroundColor Red
            Write-Host ($errorText | Out-String) -ForegroundColor Magenta
            Write-Host "Please remove the event definitions from the XAML and try again" -ForegroundColor Red
            Exit
        }

        # Create and Load the XAML Reader to create the WPF Form
        $xamlReader = (New-Object System.Xml.XmlNodeReader $newXaml)
        $uiForm = [System.Windows.Markup.XamlReader]::Load($xamlReader)

        # Create a PowerShell Variable $WPF_ for each Form Object defined in the XAML where the 
        #   element name begin with "_" as in Name="_MyObj" or x:Name="_MyObj"
        # For the most part, this allows the Code-Behind to directly access the element without needing to find it first
        $FormObjects = $newXaml.SelectNodes("//*[@Name]").Where({ $_.Name -like "_*" }).ForEach({
            $element = $uiForm.FindName($_.Name)
            If ($element) {
                Write-Output ("WPF$($_.Name)")
                Set-Variable -Name "WPF$($_.Name)" -Value $element -Force -Scope Global -Visibility Public
            } Else {
                Write-Warning "Unable to locate $($_.Name) element name"
            }
        })
        Write-Host "`nList of `$WPF_ named Variables of XAML Elements`n$($FormObjects | Out-String)`n" -ForegroundColor Green
        $uiForm
    } Catch {
        Write-Error ($_ | Format-List | Out-String)
    }
}
