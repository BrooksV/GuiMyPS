# WPF Usage Guide

## Table of Contents

- [Introduction](#introduction)
- [Prerequisites](#prerequisites)
- [Setting Up the Environment](#setting-up-the-environment)
- [Creating the WPF Application](#creating-the-wpf-application)
  - [XAML Layout](#xaml-layout)
  - [PowerShell Code-Behind Example](#powershell-code-behind-example)
- [Binding Data to the DataGrid](#binding-data-to-the-datagrid)
  - [Creating the SqlQueryDataSet Instance](#creating-the-sqlquerydataset-instance)
  - [Executing the Query](#executing-the-query)
  - [Binding the DataTable to the DataGrid](#binding-the-datatable-to-the-datagrid)
- [Handling CRUD Operations](#handling-crud-operations)
  - [Adding New Rows](#adding-new-rows)
  - [Updating Existing Rows](#updating-existing-rows)
  - [Deleting Rows](#deleting-rows)
  - [Saving Changes](#saving-changes)
- [Advanced Topics](#advanced-topics)
  - [Handling Complex Queries](#handling-complex-queries)
  - [Using Multiple DataGrids](#using-multiple-datagrids)
  - [Customizing DataGrid Columns](#customizing-datagrid-columns)
- [Troubleshooting](#troubleshooting)
- [Conclusion](#conclusion)

## Introduction

This guide provides detailed instructions on how to use the `SqlQueryClass` module to bind the `SqlQueryTable` class result as a `DataTable` to a WPF `DataGrid` component. This allows for seamless integration of SQL query results into a WPF application, enabling data manipulation and display.

## Prerequisites

Before you begin, ensure you have the following:

- PowerShell Version 5.1 or higher
- A text or code editor of you choice
- - PowerShell scripts and even XAML can created and edited with any text or code editor
- - PowerShell_ISE comes with most every Windows OS
- - VS Code or Visual Studio are also very good for serious development
- - - Helps with cloning repository and PS Code development
- - - Visual Studio comes with Blend which is a WPF GUI based editor
- Git for version control
- Basic knowledge of WPF and XAML
- SQL Server or SQL Express installed
- The `SqlQueryClass` module installed
- The example code below uses tests configuration and sample SQL Express database from [GitHub `SqlQueryClass` repository](https://github.com/BrooksV/SqlQueryClass)
- - Clone the Repository (using git commands, VS Code, or Visual Studio) to C:\Git folder to maintain compatibility with the sample and test data
- - Can also be done downloading the repository as a zip file

## Setting Up the Environment

1. Install the `SqlQueryClass` module if you haven't already:

    ```powershell
    Install-Module -Name SqlQueryClass -Repository PSGallery -Scope CurrentUser
    ```

2. Create a new PowerShell project in VS Code or Visual Studio.

- Visual Studio's WPF project is a .Net C# project and is great for creating the XAML file that can be used with PowerShell or with C#.
- There are a few differences in the XAML such as PowerShell does not support Click and other OnEvent handlers and need to be added in Code-Behind.
- The advantage is it's PowerShell which is much easier to learn, use, and edit than a traditional .Net C# application development, test, and release process.
- PowerShell WPF is great for adding simple to very complex GUI interfaces and applications for your scripts.

## Creating the WPF Application

### XAML Layout

Define the layout of your WPF application in the `MainWindow.xaml` file. Add a `DataGrid` component to display the data.

The [PowerShell Code-Behind Example](#powershell-code-behind-example) uses the XAML in-line to avoid having to read from disk.

```xml
<Window x:Class="SqlQueryWpfApp.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SQL Query DataGrid" Height="450" Width="800">
    <Grid>
        <DataGrid x:Name="_dataGrid" AutoGenerateColumns="True" />
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="10">
            <Button x:Name="_saveButton" Content="Save" Width="75" Margin="5"/>
            <Button x:Name="quitButton" Content="Quit" Width="75" Margin="5"/>
        </StackPanel>
    </Grid>
</Window>
```

### PowerShell Code-Behind Example

Code-Behind is a PowerShell example script `.\tests\Test-WpfSqlQueryClassExample.ps1` to demonstrate the usage of the GuiMyPS and SqlQueryClass modules in a WPF application and how to bind data to the `DataGrid`.

```powershell
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
```

## Binding Data to the DataGrid

### Creating the SqlQueryDataSet Instance

Create an instance of the `SqlQueryDataSet` class and configure the necessary properties such as `SQLServer`, `Database`, or `ConnectionString`.

```powershell
# Either one of the following is needed. 1st one creates a SqlConnectonString, 2nd one uses the $ConnectionString value
# $syncHash.UI.SqlResults = New-SqlQueryDataSet -SQLServer $SqlServer -Database $Database
$syncHash.UI.SqlResults = New-SqlQueryDataSet -SQLServer -ConnectionString $ConnectionString
```

### Executing the Query

Use the `ExecuteQuery` method to execute the SQL query and retrieve the results.

```powershell
$syncHash.UI.SqlResults.ExecuteQuery("SELECT * FROM [dbo].[SqlQuery]")
$dataGrid.ItemsSource = $syncHash.UI.SqlResults.Tables[0].Result.DefaultView
```

Explanation:

Class [SqlQueryDataSet] has these properties: SQLServer, Database, ConnectionTimeout, CommandTimeout, ConnectionString, SQLConnection, TableIndex, Tables, TableNames, DisplayResults, KeepAlive

- `$syncHash.UI.SqlResults` is the instance of [SqlQueryDataSet] created and returned by the New-SqlQueryDataSet function
- `$syncHash.UI.SqlResults.Tables` is the collection of [SqlQueryTable], one for each unique query added or executed by the class
- `$syncHash.UI.SqlResults.Tables[0]` is the first instance of [SqlQueryTable] which is the Query just created by ExecuteQuery() method
- `$syncHash.UI.SqlResults.Tables[0].Result` is the ExecuteQuery() results returned as a [System.Data.DataTable]
- - Execute methods always saves data to Tables[0].Result
- `$syncHash.UI.SqlResults.Tables[0].Result.DefaultView` is the [System.Data.DataView] required for binding to the WPF XAML DataGrid component

Other `$syncHash.UI.SqlResults` properties:
- TableIndex -- `[int]` value that represents the currently selected index used in `Tables[TableIndex]`
- TableNames -- HashTable of Unique Query Table identifiers and the index in the Table[] collection
- - Used to lookup index by TableName, Example: `$syncHash.UI.SqlResults.TableNames['DBTable']` returns 0
- DisplayResults -- Helpful when wanting to see the results from the PS Pipeline
- - True sends the results to the PS Pipeline after its saved to `$syncHash.UI.SqlResults.Tables[0].Result`
- - False only saves Result to `$syncHash.UI.SqlResults.Tables[0].Result`

Class [SqlQueryTable] has these properties: TableIndex, TableName, Query, SQLCommand, SqlDataAdapter, ResultType, Result, isDirty, QueryFile, Parent

For WPF data binding to work both ways, properties (Result, SqlDataAdapter, and SQLCommand) need to be persistent to performing CRUD actions. A WPF application might involve many tables and datasets for controls like comboboxes, lists, datagrids, treeviews, ect.. This is why [SqlQueryTable] was needed for every unique query dataset.

See [API Guide and Class Documentation](api.guide.md) for detailed information.

### Binding the DataTable to the DataGrid

Bind the `DataTable` result to the `DataGrid` by setting the `ItemsSource` property to the `DefaultView` of the `DataTable`.

## Handling CRUD Operations

### Adding New Rows

To add new rows to the `DataGrid`, modify the `DataTable` and call the `SaveChanges` method on the `SqlQueryDataSet` instance.

### Updating Existing Rows

To update existing rows, modify the `DataTable` and call the `SaveChanges` method.

### Deleting Rows

To delete rows, remove them from the `DataTable` and call the `SaveChanges` method.

### Saving Changes

Call the `SaveChanges` method on the `SqlQueryDataSet` instance to persist changes to the database.

## Advanced Topics

### Handling Complex Queries

Learn how to handle complex queries and multiple result sets.

### Using Multiple DataGrids

Bind multiple `DataGrids` to different `DataTables` within the same `SqlQueryDataSet` instance.

### Customizing DataGrid Columns

Customize the columns of the `DataGrid` to display specific data and apply formatting.

## Troubleshooting

Common issues and solutions when working with the `SqlQueryClass` module and WPF `DataGrid`.

## Conclusion

By following this guide, you can effectively bind SQL query results to a WPF `DataGrid` using the `SqlQueryClass` module. This enables seamless data manipulation and display within your WPF applications.
