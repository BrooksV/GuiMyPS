# Developer and Contributor Guide

## Table of Contents

- [Setup](#setup)
- [Source Files used in the Module](#source-files-used-in-the-module)
- [Module Build Process](#module-build-process)
- [Publishing `GuiMyPS` Module to GitHub](#publishing-GuiMyPS-module-to-github)
- [Code Review and Feedback](#code-review-and-feedback)
- [Merge the Pull Request](#merge-the-pull-request)
- [Cleanup](#cleanup)
- [Publishing `GuiMyPS` Module to PowerShell Gallery](#publishing-GuiMyPS-module-to-powershell-gallery)

## Setup

- Uses SQL Express but should work with other SQL Databases with proper connection strings and credentials
- Requires VS Code
- For Contributors, Fork the [GuiMyPS](https://github.com/BrooksV/GuiMyPS) repository
- Clone the repository or fork to local pc. I like using c:\git as my local repository folder. Subfolder $ProjectName will be created with the GitHub repository contents
- Install [Manjunath Beli's ModuleTools](https://github.com/belibug/ModuleTools) module as the module build process uses ModuleTools

  ```powershell
  Find-Module -Name ModuleTools | Install-Module -Scope CurrentUser -Verbose
  ```

- Note that a sample SQL Express database file (.\tests\TestDatabase1.mdf) is included for pester tests. The database configuration is set in .\tests\TestDatabase1.parameters.psd1

## Source Files used in the Module

- Public functions that are exported, are separate files in the .\src\public folder.
- Private functions that are local to the Module, are separate files in the .\src\private folder.
- Resources are files and folders in the .\src\resources folder that needs to be included with the Manifest and Module

## Module Build Process

1. Create a local branch for your changes.

- - Use descriptive name that reflects the type of changes for branch for example features/database-table-access

   ```powershell
   git checkout -b features/database-table-access
   ```

1. Update the build version using Update-MTModuleVersion (Find-Module -Name ModuleTools).
1. Commit your changes to the branch.
1. Run the Pester Tests using Invoke-MTTest (Find-Module -Name ModuleTools).
1. Build the Module output using Invoke-MTBuild -Verbose (Find-Module -Name ModuleTools).

   ```powershell
   Invoke-MTBuild -Verbose
   ```

- - Outputs to the .\dist\GuiMyPS folder
- - Combines the file contents of the files in Public and Private folder into .\dist\GuiMyPS\GuiMyPS.psd1 and exports the Public Functions
- - Generates the .\dist\GuiMyPS\GuiMyPS.psd1 Manifest file from the settings in .\project.json
- - Resources (.\src\resources) folder content is copied to .\dist\GuiMyPS folder
- Run the Pester Teats using Invoke-MTTest (Find-Module -Name ModuleTools)

1. Make corrections, repeat the build process.
1. For Contributors:
   - Create an Issue if one does not exist that addresses the proposed changes.
   - Upstream your branch.
   - Create a Pull request.

## Publishing `GuiMyPS` Module to GitHub

1. Stage and Commit Your Changes

   ```powershell
   git add .
   git commit -m "Implemented database and table access functions"
   ```

1. Update remote repository with branch changes

   ```powershell
    # List status of remote repository
    git branch -r
    # Create Branch on remote repository if needed
    # git push --set-upstream origin features/database-table-access
    # Push branch changes to remote branch in repository
    git push origin features/database-table-access
    ```

1. Create a Pull Request on remote repository
   - Go to [GuiMyPS GitHub repository](https://github.com/BrooksV/GuiMyPS)
   - Click on "Compare & pull request" for your branch
   - Provide a meaningful title and description for the PR
   - Select the base branch (main) to merge into
   - Click "Create pull request"

## Code Review and Feedback

- Engage with Repository Owner or collaborators to review the PR
- Address any feedback or requested changes by making additional commits to your branch and pushing them to the remote branch
- Ensure the PR passes any automated tests or checks

## Merge the Pull Request

- Once the PR is approved and all checks pass, you can merge it into the main branch
- You can either use the "Merge pull request" button on GitHub or merge it locally and push the changes

## Cleanup

- After merging, you can delete the feature branch from the remote repository to keep it clean

```powershell
git push origin --delete features/database-table-access
```

- Optionally, delete the local branch

```powershell
git branch -d features/database-table-access
```

These steps will ensure your changes are integrated into the main branch and your repository remains organized.

## Publishing `GuiMyPS` Module to PowerShell Gallery

```powershell
$data = Get-MTProjectInfo
$ApiKey = "your-api-key-here"
Publish-Module -Path $data.OutputModuleDir -NuGetApiKey $ApiKey -Repository PSGallery
```
