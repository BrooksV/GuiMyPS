<#
.SYNOPSIS
Verifies that the required modules are installed and imported.

.DESCRIPTION
This function checks if a specified module is installed. 
If not, it installs the module from the PSGallery repository and imports it. 
If the module is already installed, it simply imports the module.

.PARAMETER ModuleName
The name of the module to verify, install, and import.

.EXAMPLE
Initialize-Module -ModuleName "SqlQueryClass"

.EXAMPLE
Initialize-Module -ModuleName "GuiMyPS"

.EXAMPLE
Initialize-Module -ModuleName "ImportExcel"

.NOTES
Author: Brooks Vaughn
Date: "22 February 2025"
#>

Function Initialize-Module {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$ModuleName
    )

    # Check if the module is already installed
    $module = Get-Module -ListAvailable -Name $ModuleName
    If ($null -eq $module) {
        # Module is not installed, install it
        Write-Host "Module '$ModuleName' is not installed. Installing..."
        Install-Module -Name $ModuleName -Repository PSGallery -Scope CurrentUser
    
        # Import the newly installed module
        Write-Host "Importing module '$ModuleName'..."
        Import-Module -Name $ModuleName
    } Else {
        # Module is already installed, import it
        Write-Host "Module '$ModuleName' is already installed. Importing..."
        Import-Module -Name $ModuleName
    }

    # Verify the module is imported
    If (Get-Module -Name $ModuleName) {
        Write-Host "Module '$ModuleName' has been successfully imported."
    } Else {
        Write-Host "Failed to import module '$ModuleName'."
    }
}
