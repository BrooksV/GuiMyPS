

Import-Module C:\Git\GuiMyPS\dist\GuiMyPS\GuiMyPS.psd1 -Verbose -Force
# Remove-Module -Name GuiMyPS -Force -Verbose
# Import-Module -Name GuiMyPS -Force -Verbose

# . "C:\Git\GuiMyPS\src\public\New-XamlWindow.ps1"

$Error.Clear()

<# Example with Events defined # >
$inputXML = @"
<Window x:Class="WpfApp.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Event Handling Example" Height="450" Width="800">
    <Grid>
        <StackPanel Margin="10">
            <Button x:Name="myButton" Content="Click Me" Width="100" Click="myButton_Click" />
            <TextBox x:Name="myTextBox" Width="300" TextChanged="myTextBox_TextChanged" />
            <ComboBox x:Name="myComboBox" Width="150" SelectionChanged="myComboBox_SelectionChanged">
                <ComboBoxItem Content="Option 1" />
                <ComboBoxItem Content="Option 2" />
                <ComboBoxItem Content="Option 3" />
            </ComboBox>
            <CheckBox x:Name="myCheckBox" Content="Check Me" Checked="myCheckBox_Checked" Unchecked="myCheckBox_Unchecked" />
            <Slider x:Name="mySlider" Width="200" ValueChanged="mySlider_ValueChanged" />
        </StackPanel>
    </Grid>
</Window>
"@
#>

<# Example with "_" names variables #>
$inputXML = @"
<Window x:Class="ScanMyBills.MainWindow"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:system="clr-namespace:System;assembly=mscorlib"
    xmlns:local="clr-namespace:ScanMyBills"
    mc:Ignorable="d"
    Title="SQL Query DataGrid" 
    Height="800" Width="1200" Topmost="False" 
    ResizeMode="CanResizeWithGrip" ShowInTaskbar = "True"
    WindowStartupLocation = "CenterScreen"
    x:Name="MainForm"
    FocusManager.FocusedElement="{Binding ElementName=_scriptView}"
    Background="AliceBlue" UseLayoutRounding="True"
    >
    <Grid>
        <DataGrid x:Name='_dataGrid' AutoGenerateColumns='True' />
        <StackPanel Orientation='Horizontal' HorizontalAlignment='Right' VerticalAlignment='Bottom' Margin='10'>
            <Button x:Name='_saveButton' Content='Save' Width='75' Margin='5' />
            <Button x:Name='quitButton' Content='Quit' Width='75' Margin='5' />
        </StackPanel>
    </Grid>
</Window>
"@
#>

<# Test for XAML String # >
try {
    # $form1 = (New-XamlWindow -xaml $inputXML) 
    $xaml = New-XamlWindow -xaml $inputXML
    $xamlReader = (New-Object System.Xml.XmlNodeReader $xaml)
    $form1 = [System.Windows.Markup.XamlReader]::Load($xamlReader)
    $form1.ShowDialog()
} catch {
    Write-Warning ($_ | Format-List | Out-String)
}
#>

$handler_Button_Click = {
    Param ([object]$theSender, [System.EventArgs]$e)
    Write-Host ("`$handler_Button_Click() Item clicked: {0}" -f $theSender.Name)
    Switch -Regex ($theSender.Name) {
            '^_saveButton$' {
                Break
            }
            '^quitButton$' {
                $rootElement = Find-RootElement -Element $theSender
                If ($rootElement) {
                    $rootElement.Close()
                }
                Break
            }
        default {
            Write-Host ("{0}: {1}({2})" -f $theSender.Name, $e.OriginalSource.Name, $e.OriginalSource.ToString())
        }
    }
}

$handler_MenuItem_Click = {
    Param ([object]$theSender, [System.EventArgs]$e)
    Write-Host ("`$handler_MenuItem_Click() Item clicked: {0}" -f $theSender.Name)
    Switch -Regex ($theSender.Name) {
        '^_toggleEditMode$' {
            Break
        }
        '^mnuMExport$' {
            Break
        }
        '^mnuMExportCopyToClipboard$' {
            Break
        }
        '^mnuMExportAsCSV$' {
            Break
        }
        '^mnuMExportAsExcel$' {
            Break
        }
        '^mnuMImportFromExcel$' {
            Break
        }
        '^mnuMAdd$' {
            Break
        }
        '^mnuMCancel$' {
            Break
        }
        '^mnuMSave$' {
            Break
        }
        '^mnuMDelete$' {
            Break
        }
        '^mnuMFirst$' {
            Break
        }
        '^mnuMPrevious$' {
            Break
        }
        '^mnuMNext$' {
            Break
        }
        '^mnuMLast$' {
            Break
        }
        '^mnuDAdd$' {
            Break
        }
        '^mnuDCancel$' {
            Break
        }
        '^mnuDSave$' {
            Break
        }
        '^mnuDDelete$' {
            Break
        }
        default {
            Write-Host ("{0}: {1}({2})" -f $theSender.Name, $e.OriginalSource.Name, $e.OriginalSource.ToString())
        }
    }
}

<# Test for XAML String #>
try {
    $form1 = (New-XamlWindow -xaml $inputXML)
    # Add-ClickToEveryButton -Element $form1 -ClickHandler $handler_button_Click
    $elements = @()
    $elements += Find-EveryControl -Element $form1 -ControlType 'System.Windows.Controls.Button'
    # Build-HandlerCode -Elements $elements -ControlType System.Windows.Controls.Button
    $elements.ForEach({$_.Element.Add_Click($handler_Button_Click)})
    $form1.ShowDialog()
} catch {
    Write-Warning ($_ | Format-List | Out-String)
}
#>

$Error
Break

# Test for [System.Xml.XmlDocument]
try {
    $xaml = [xml]$inputXML
    $form2 = New-XamlWindow -xaml $xaml
    Add-ClickToEveryButton -Element $form2 -ClickHandler $handler_button_Click
    $form2.ShowDialog()
} catch {
    Write-Warning ($_ | Format-List | Out-String)
}

# Test for File Path to XAML file
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

$Error
