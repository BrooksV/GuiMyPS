# GuiMyPS

Module helps in creating and running WPF GUI based PowerShell Applications.

Original List of Potential Module Names to use:

PsWpfHelper
WpfHelperForPS
PsWpfUtils
WpfUtilsForPS
PsWpfLib
PsXaml
PsXamlHelper
GuiMyPS - Collections of XAML & WPF Helper Functions to simplify creation of GUI based PowerShell Applications. 

Decided on using:
GuiMyPS - Module helps in creating and running WPF GUI based PowerShell Applications.

## Coding Style and Copilot instructions

This is my PowerShell coding style requirements:
Brace Styles use "K&R (Kernighan and Ritchie) style"
Variables use CamelCase
Operators and logical conditions use lowercase
Statement Keywords use PascalCase
Parameter names use PascalCase
Functions should include Helpful Metadata and documentation sections which includes .NOTES as the last section before the Function Declaration
NOTES section should set the Author: Brooks Vaughn and set Date to Date: to the current date as a string literal

Update the following function to comply with my coding style. Before presenting the results, fix the keywords in the result that do not comply with the coding style

Update the function to comply with my coding style. Before presenting the results, fix the keywords in the result that do not comply with the coding style

## `GuiMyPS` Module and Status Details

Name          | Version | PS Compatibility | Project Uri (GitHub)
------------- | ------- | ---------------- | ------------------------------------------------------------------------------------
GuiMyPS | 0.0.1   | 5.1              | [https://github.com/BrooksV/GuiMyPS](https://github.com/BrooksV/GuiMyPS)

[PSGalleryLink]: https://www.powershellgallery.com/packages/GuiMyPS/
[BadgeIOCount]: https://img.shields.io/powershellgallery/dt/GuiMyPS.svg?label=downoads%20GuiMyPS%40PSGallery
[WorkFlowStatus]: https://img.shields.io/github/actions/workflow/status/BrooksV/GuiMyPS/tests.yml?label=tests.yml%20build

![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/GuiMyPS.svg)
[![maintainer](https://img.shields.io/badge/maintainer-BrooksV-orange)](https://github.com/BrooksV)
[![License](https://img.shields.io/github/license/BrooksV/GuiMyPS)](https://github.com/BrooksV/GuiMyPS/blob/main/LICENSE)
[![contributors](https://img.shields.io/github/contributors/BrooksV/GuiMyPS.svg)](https://github.com/BrooksV/GuiMyPS/graphs/contributors/)
[![last-commit](https://img.shields.io/github/last-commit/BroksV/GuiMyPS.svg)](https://github.com/BrooksV/GuiMyPS/commits/)
[![issues](https://img.shields.io/github/issues/BrooksV/GuiMyPS.svg)](https://github.com/BrooksV/GuiMyPS/issues/)
[![issues-closed](https://img.shields.io/github/issues-closed/BrooksV/GuiMyPS.svg)](https://github.com/BrooksV/GuiMyPS/issues?q=is%3Aissue+is%3Aclosed)
[![GitHub stars](https://img.shields.io/github/stars/BrooksV/GuiMyPS.svg)](https://github.com/BrooksV/GuiMyPS/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/BrooksV/GuiMyPS.svg)](https://github.com/BrooksV/GuiMyPS/network/members)
[![GitHub pull requests](https://img.shields.io/github/issues-pr/BrooksV/GuiMyPS.svg)](https://github.com/BrooksV/GuiMyPS/pulls)

### Build and Release Statistics

[![GuiMyPS@PowerShell Gallery][BadgeIOCount]][PSGalleryLink]
![WorkFlow Status][WorkFlowStatus]
![Build Status](https://img.shields.io/github/actions/workflow/status/BrooksV/GuiMyPS/ci.yml?label=ci.yml%20build)

![Version](https://img.shields.io/github/v/release/BrooksV/GuiMyPS.svg?label=version)
![GitHub All Releases](https://img.shields.io/github/downloads/BrooksV/GuiMyPS/total.svg?label=release%20dl%20all%40GitHub)
![GitHub release (latest by date)](https://img.shields.io/github/downloads/BrooksV/GuiMyPS/latest/total.svg?label=release%20dl%20by%20date%40GitHub)
![Downloads](https://img.shields.io/github/downloads/BrooksV/GuiMyPS/total.svg?label=total%20release%20dl%40GitHub)

### Related Links

- [LicenseUri](https://github.com/BrooksV/GuiMyPS/blob/main/LICENSE)
- [ProjectUri](https://github.com/BrooksV/GuiMyPS)

### Module Tags / Keywords

[PowerShell], [WPF], [XAML], [GUI]

## Installation

```powershell
Install-Module -Name GuiMyPS -Repository PSGallery -Scope CurrentUser
```

To load a local build of the module, use Import-Module as follows:

```powershell
Import-Module -Name "C:\Git\GuiMyPS\dist\GuiMyPS\GuiMyPS.psd1" -Force -verbose
```

### Requirements

- Tested with PowerShell 5.1 and 7.5x
- No known dependencies for usage
- VS Code and clone [Brooks Vaughn's GuiMyPS](https://github.com/BrooksV/GuiMyPS) Repository
- Module build process uses [Manjunath Beli's](https://github.com/belibug) [ModuleTools](https://github.com/belibug) module.
- Test scripts requires the Pester module and SQL Express
- Includes sample SQL Express database file used in test scripts

### Features

- Includes helper functions

## Usage

The GuiMyPS Module was developed to support the loading of WPF (Windows Presentation Framework) XAML elements, locating elements, adding Click and other Event Handlers, and miscellaneous helper functions.

### New-XamlWindow Helper Function

Used to prepare and load a XAML string, filepath, or XML Document returning the WPF [System.Windows.Window] form object

```powershell
[System.Windows.Window]New-XamlWindow [-xaml] <Object> [-NoXRemoval] [-WhatIf] [-Confirm] [<CommonParameters>]
```

Event handlers and code-behind logic is added to the Form Object after it's creation and before the .ShowDialog() method is called.

#### New-XamlWindow Help Links

- For usage examples
- - [Test-Usage-01.ps1](.\tests\Test-Usage-01.ps1)
- - [Test-WpfSqlQueryClassExample.ps1](.\tests\Test-WpfSqlQueryClassExample.ps1)
- For examples, type: "Get-Help New-XamlWindow -Examples"
- For detailed information, type: "Get-Help New-XamlWindow -Detailed"
- For technical information, type: "Get-Help New-XamlWindow -Full"
- See [WPF Guide](wpf.guide.md) for detailed information about WPF.

### Example 1: XAML as String

```powershell
$form = New-XamlWindow -xaml $xamlString
```

### Example 2: XAML as FilePath

```powershell
$form = New-XamlWindow -xaml MainWindows.xaml
```

### Example 3: XAML as XML Document

```powershell
$form = New-XamlWindow -xaml [xml]$xaml
```

## How to Contribute

1. Fork the repository.
2. Create a local branch for your changes.
3. Make your changes and commit them.
4. Push your changes to your fork.
5. Create a pull request.

For detailed information on "How to Contribute", set up the development environment, and more, please refer to the [Developer and Contributor Guide](contributor.guide.md).

This includes details on:

- [Setup](contributor.guide.md#setup)
- [Source Files used in the Module](contributor.guide.md#source-files-used-in-the-module)
- [Module Build Process](contributor.guide.md#module-build-process)
- [Publishing `GuiMyPS` Module to GitHub](contributor.guide.md#publishing-GuiMyPS-module-to-github)
- [Code Review and Feedback](contributor.guide.md#code-review-and-feedback)
- [Merge the Pull Request](contributor.guide.md#merge-the-pull-request)
- [Cleanup](contributor.guide.md#cleanup)
- [Publishing `GuiMyPS` Module to PowerShell Gallery](contributor.guide.md#publishing-GuiMyPS-module-to-powershell-gallery)

### ToDo

- [ ] Include Pester Tests
- [ ] Add comment-based help sections to all CmdLets and Functions
- [ ] Seek peer review and comments
- [ ] Integrate feedback
- [ ] Improve Documentation
- [ ] Create a comprehensive fully functional WPF PowerShell Application

## Module Exported Functions

```powershell
Get-Command -Module "GuiMyPS" -Syntax

- Add-ClickToEveryButton [-Element] <Object> [[-ClickHandler] <RoutedEventHandler>] [<CommonParameters>]
- Add-ClickToEveryMenuItem [-MenuObj] <Object> [-Handler] <scriptblock> [<CommonParameters>]
- Add-EventsToEveryCheckBox [-Element] <Object> [[-ClickHandler] <RoutedEventHandler>] [[-CheckedHandler] <RoutedEventHandler>] [[-UncheckedHandler] <RoutedEventHandler>] [[-PreviewMouseUpHandler] <MouseButtonEventHandler>] [-PreventSelectionScrolling] [<CommonParameters>]
- Build-HandlerCode [-Elements] <Object[]> [-ControlType] <string> [<CommonParameters>]
- Find-EveryControl [-ControlType] <Object> [-Element] <Object> [-ExcludeElement] [-UseVisualTreeHelper] [-IncludeAll] [<CommonParameters>]
- Find-RootElement [-Element] <DependencyObject> [<CommonParameters>]
- Format-XML [-xml] <xml> [[-indent] <Object>] [-FormatAttributes] [-IncludeXmlDeclaration] [<CommonParameters>]
- Get-ControlContent [-Element] <Object> [<CommonParameters>]
- Get-FormVariable [<CommonParameters>]
- Get-ObjectPropertyDetail [-InputObject] <Object> [<CommonParameters>]
- New-XamlWindow [-xaml] <Object> [-NoXRemoval] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Details

- Add-ClickToEveryButton -- Adds a Click Event Handler to all buttons nested with the starting element
- Add-ClickToEveryMenuItem -- Adds a Click Event Handler to all MenuItems nested with the starting Menu Object / element
- Add-EventsToEveryCheckBox -- Adds a various Event Handlers to all CheckBoxes nested with the starting element
- Build-HandlerCode -- Generates the code for an Event Handler from a list of Elements
- Find-EveryControl -- Searches for and creates a list of elements found of a specific Control Type with options to list all controls
- Find-RootElement -- Transverse up the parent object tree looking for the root element which is the WPF form object
- Format-XML -- XML pretty formatter / printer
- Get-ControlContent -- An attempt to extract the display text for a control
- Get-FormVariable -- Lists the $WPF_* global variables that were created by New-XamlWindow when parsing the XAML x:Name="_elementName" names
- Get-ObjectPropertyDetail -- Dumpts the Property, Value, and Type of any PowerShell object
- New-XamlWindow -- Main CmdLet that prepares and loads a WPF XAML string, filepath, or XML Document returning a WPF Form Object

## Folder Structure and Build Management

The folder structure of the GuiMyPS module is based on best practices for PowerShell module development and was initially created using [Manjunath Beli's](https://github.com/belibug) [ModuleTools](https://github.com/belibug) module. Check out his [Blog article](https://blog.belibug.com/post/ps-modulebuild) that explains the core concepts of ModuleTools.

Install [Manjunath Beli's ModuleTools](https://github.com/belibug/ModuleTools) module as the module build process uses ModuleTools

```powershell
Find-Module -Name ModuleTools | Install-Module -Scope CurrentUser -Verbose
```

The the following ModuleTools CmdLets used in the build and maintenance process. They need to be executed from project root:

- Get-MTProjectInfo -- returns HashTable of project configuration which can be used in pester tests or for general troubleshooting
- Update-MTModuleVersion -- Increments GuiMyPS module version by modifying the values in `project.json` or you can manually edit the json file.
- Invoke-MTBuild -- Run `Invoke-MTBuild -Verbose` to build the module. The output will be saved in the `dist` folder, ready for distribution.
- Invoke-MTTest -- Executes pester configuration (*.text.ps1) files in the `tests` folder

- To skip a test, add `-skip` in describe block of the Pester *.test.ps1 file to skip.

### Folder and Files

```powershell
.\GuiMyPS
|   .gitignore
|   CODE_OF_CONDUCT.md
|   contributor.guide.md
|   git.cheatsheet.md.ps1
|   GitHub_Action_Docs.md
|   LICENSE
|   project.json
|   ps.readinglist.md
|   README.md
|   wpf.guide.md
|
+---.github
|   \---workflows
|           tests.yml
|
+---archive
+---dist
|   \---GuiMyPS
|           about_GuiMyPS.help.txt
|           GuiMyPS.psd1
|           GuiMyPS.psm1
|
+---src
|   +---private
|   +---public
|   |       Add-ClickToEveryButton.ps1
|   |       Add-ClickToEveryMenuItem.ps1
|   |       Add-EventsToEveryCheckBox.ps1
|   |       Build-HandlerCode.ps1
|   |       Find-EveryControl.ps1
|   |       Find-RootElement.ps1
|   |       Format-XML.ps1
|   |       Get-ControlContent.ps1
|   |       Get-FormVariable.ps1
|   |       Get-ObjectPropertyDetail.ps1
|   |       New-XamlWindow.ps1
|   |
|   \---resources
|           about_GuiMyPS.help.txt
|
\---tests
        Module.Tests.ps1
        OutputFiles.Tests.ps1
        ScriptAnalyzer.Tests.ps1
        Test-Usage-01.ps1
        Test-WpfSqlQueryClassExample.ps1
        TestDatabase1.mdf
        TestDatabase1.parameters.psd1
        TestDatabase1_log.ldf
```

All files and folders in the `src` folder, will be published Module.

All other folder and files in the `.\GuiMyPS` folder will resides in the [GitHub GuiMyPS Repository](https://github.com/BrooksV/GuiMyPS) except those excluded by inclusion in the `.\GuiMyPS\.gitignore` file.

### Project JSON File

The `project.json` file contains all the important details about your module, is used during the module build process, and helps to generate the GuiMyPS.psd1 manifest.

### Root Level and Other Files

- .gitignore -- List of file, folder, and wildcard specifications to ignore when publishing to GitHub repository
- CODE_OF_CONDUCT.md -- Standard GitHub code of conduct and standards
- GitHub_Action_Docs.md -- How to add GitHub Action WorkFlows to automate CI/CD (Continuous Integration/Continuous Deployment)
- LICENSE -- MIT License notice and copyright
- project.json -- ModuleTools project configuration file used to build the `GuiMyPS` module
- README.md -- Documentation (this) file for the `GuiMyPS` module
- .vscode\settings.json -- VS Code settings used during `GuiMyPS` module development
- *.guide.md -- various guides such as contributor, wpf

### archive Folder

`.\GuiMyPS\archive` is not used in this project. Its a temporary place / BitBucket to hold code snippets and files during development and is not part of the build.

### Dist (build output) Folder

Generated module is stored in `dist\GuiMyPS` folder, you can easily import it or publish it to PowerShell Gallery or repository.

### Src Folder

- All functions in the `public` folder are exported during the module build.
- All functions in the `private` folder are accessible internally within the module but are not exposed outside the module.
- All files and folder contained in the `resources` folder will be published to the `dist\GuiMyPS` folder.

### Tests Folder

If you want to run any `pester` tests, keep them in `tests` folder and named *.test.ps1.

Run `Invoke-MTTest` to execute the tests.

- .\src\public\New-XamlWindow.tests.ps1 -- Full set of usage example Tests. Good Resource for usage examples
- .\tests\Module.Tests.ps1 -- General Module Control to verify the module imports correctly
- .\tests\OutputFiles.Tests.ps1 -- Module and Manifest testing to verify output files are readable
- .\tests\ScriptAnalyzer.Tests.ps1 -- Code Quality Checks to verify PowerShell syntax and best practices
- .\tests\TestDatabase1.parameters.psd1 -- PowerShell Data File of configuration settings used in Test-WpfSqlQueryClassExample.ps1
- .\tests\TestDatabase1.mdf -- Sample SQL Express Database File with samples data used in Test-WpfSqlQueryClassExample.ps1
- .\tests\TestDatabase1_log.ldf -- Created when using TestDatabase1.mdf

## Join the Conversation

We encourage you to participate in our [Discussions](https://github.com/BrooksV/GuiMyPS/discussions) section! Whether you have questions, ideas, or just want to chat with other users, Discussions is the place to be. Your feedback and contributions are valuable to us!

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgements

- [Manjunath Beli](https://github.com/belibug) for the [ModuleTools](https://github.com/belibug/ModuleTools) module used in the build process.
- [Brooks Vaughn](https://github.com/BrooksV) for maintaining the GuiMyPS module.

## Contact

For support, inquiries, or feedback, contact Brooks Vaughn at [BrooksV](https://github.com/BrooksV) through one of the following methods:

- **GitHub Issues**: [Open an issue](https://github.com/BrooksV/GuiMyPS/issues)
- **GitHub Discussions**: [Start a discussion](https://github.com/BrooksV/GuiMyPS/discussions)
