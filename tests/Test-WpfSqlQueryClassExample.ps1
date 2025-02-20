Function Initialize-Module {
    Param (
        [string]$moduleName
    )
    # Check if the module is already installed
    $module = Get-Module -ListAvailable -Name $moduleName
    if ($null -eq $module) {
        # Module is not installed, install it
        Write-Output "Module '$moduleName' is not installed. Installing..."
        Install-Module -Name $moduleName -Repository PSGallery -Scope CurrentUser
    
        # Import the newly installed module
        Write-Output "Importing module '$moduleName'..."
        Import-Module -Name $moduleName
    } else {
        # Module is already installed, import it
        Write-Output "Module '$moduleName' is already installed. Importing..."
        Import-Module -Name $moduleName
    }
    # Verify the module is imported
    if (Get-Module -Name $moduleName) {
        Write-Output "Module '$moduleName' has been successfully imported."
    } else {
        Write-Output "Failed to import module '$moduleName'."
    }
}    
# Verify required modules are installed and imported
Initialize-Module -moduleName "SqlQueryClass"
Initialize-Module -moduleName "GuiMyPS"
# Import-Module C:\Git\GuiMyPS\dist\GuiMyPS\GuiMyPS.psd1 -Verbose -Force

# Add required .Net Assemblies required for WPF
Add-Type -AssemblyName PresentationFramework

# Define variable `$syncHash` for global use to store synchronized data and provide access to object data.
$syncHash = [System.Collections.Hashtable]::Synchronized((New-Object System.Collections.Hashtable))
$syncHash.Add('UI', [PSCustomObject]@{
    SqlResults = $null
})

# Database Configuration
# Using sample database configuration data from tests `.\tests\TestDatabase1.parameters.psd1` and the SQL Express test database `.\tests\TestDatabase1.mdf`.

$SqlServer = '(localdb)\MSSQLLocalDB'
$Database = 'TestDatabase1'
$ConnectionString = "Data Source=$SqlServer;AttachDbFilename=C:\Git\SqlQueryClass\tests\TestDatabase1.mdf;Integrated Security=True"

# Use the `New-SqlQueryDataSet` function to create and initialize the `SqlQueryDataSet` instance.
$syncHash.UI.SqlResults = New-SqlQueryDataSet -SQLServer $SqlServer -Database $Database -ConnectionString $ConnectionString

$xamlString = @"
<Window x:Class="SqlQueryClass.MainWindow"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:system="clr-namespace:System;assembly=mscorlib"
    xmlns:local="clr-namespace:SqlQueryClass"
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

$handler_Button_Click = {
    Param ([object]$theSender, [System.EventArgs]$e)
    Write-Host ("`$handler_Button_Click() Item clicked: {0}" -f $theSender.Name)
    Switch -Regex ($theSender.Name) {
            '^_saveButton$' {
                $syncHash.UI.SqlResults.SaveChanges()
                [System.Windows.MessageBox]::Show("<code to support CRUD actions is needed>`n`nChanges saved successfully.", 
                    "Save", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
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

try {
    $syncHash.Form = New-XamlWindow -xaml $xamlString
    $elements = @()
    $elements += Find-EveryControl -Element $syncHash.Form -ControlType 'System.Windows.Controls.Button'
    $elements.ForEach({$_.Element.Add_Click($handler_Button_Click)})

    $syncHash.UI.SqlResults.DisplayResults = $false
    $syncHash.UI.SqlResults.ExecuteQuery("SELECT * FROM [dbo].[SqlQuery]")
    $WPF_dataGrid.ItemsSource = $syncHash.UI.SqlResults.Tables[0].Result.DefaultView

    $syncHash.Form.ShowDialog()
} catch {
    Write-Warning ($_ | Format-List | Out-String)
}
