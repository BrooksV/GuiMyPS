<#
.SYNOPSIS
Creates a unique filename by adding a sequence number or timestamp to ensure the filename is unique.

.DESCRIPTION
This function creates a unique filename based on an input filename used as a template. 
It adds a date timestamp and/or a sequence number (.01, .02,...,.99) if the new file name already exists.

.PARAMETER Name
The existing filename to make Unique

.PARAMETER NoDate
Specifies not to add a date to the filename.

.PARAMETER AddTime
Specifies to add a time to the filename.

.PARAMETER AddSeq
Specifies to add a sequence number to the filename.

.EXAMPLE
Get-UniqueFileName -Name "FileNameTemplate.ext"

.EXAMPLE
Get-UniqueFileName -Name "C:\Temp\Prod"

.EXAMPLE
Get-UniqueFileName -Name "C:\Temp\Prod" -AddSeq

.EXAMPLE
Get-UniqueFileName -Name "C:\Temp\Prod" -NoDate

.EXAMPLE
Get-UniqueFileName -Name "C:\Temp\Prod" -NoDate -AddSeq

Get-UniqueFileName -Name C:\Git\GuiMyPS\src\public\Get-UniqueFileName.ps1 -NoDate -AddSeq

Get-UniqueFileName -Name FileNameTemplate_2025-0223.ext -NoDate -AddSeq

.NOTES
Author: Brooks Vaughn
Date: "22 February 2012"
#>


Function Get-UniqueFileName {
    [CmdletBinding()]
    Param (
        [Parameter(HelpMessage="Existing FileName to Copy", ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [switch]$NoDate,
        [switch]$AddTime,
        [switch]$AddSeq
    )
    Begin {
        $NewFileName = [String]::Empty
        # Save File Ext and Get Root of File
        $PathToFile = [System.IO.Path]::GetDirectoryName($Name)
        $Ext = [System.IO.Path]::GetExtension($Name)
        $FileName = [System.IO.Path]::GetFileNameWithoutExtension($Name)

        # Strip Off Date / Time
        $FileName = ($FileName -replace "_\d{4}-\d{4}\D{1}?\d{4,6}|_\d{4}-\d{4}", "")

        # Strip off any Sequence such as .000 to .999 extension
        If ([System.IO.Path]::GetExtension($FileName) -match "\.\d{3}") {
            $FileName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
        }
    }
    Process {
        Try {
            If (-not $NoDate) {
                # Add Date Stamp
                $FileName += "_$(Get-Date -Format yyyy-MMdd)"

                # Add Time Stamp
                If ($AddTime) {
                    # Add Time Stamp to FileName
                    $FileName += "-$(Get-Date -Format HHmm)"
                }
            }

            if ([String]::IsNullOrEmpty($newFileName)) {
                $newFileName = [System.IO.Path]::Combine($pathToFile, "$FileName$($ext)")
            }            

            # Verify that New FileName does not already exist and Add Sequence # when it does
            $i = 0
            While ((Test-Path -Path $NewFileName) -or $AddSeq) {
                $AddSeq = $false
                $i++
                $Seq = "{0:D3}" -f ([int]$i)
                $NewFileName = [System.IO.Path]::Combine($PathToFile, "$FileName.$Seq$($Ext)")
            }
            $Host.UI.WriteDebugLine("Get-UniqueFileName() New FileName: ($NewFileNameItem)")
        } Catch {
            Write-Error $_
        }
    }

    End {
        $NewFileName
    }
}
