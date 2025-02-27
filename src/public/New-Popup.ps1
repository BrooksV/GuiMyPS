<#
.SYNOPSIS
New-Popup -- Display a Popup Message

.DESCRIPTION
This command uses the Wscript.Shell PopUp method to display a graphical message
box. You can customize its appearance of icons and buttons. By default the user
must click a button to dismiss but you can set a timeout value in seconds to 
automatically dismiss the popup. 

The command will write the return value of the clicked button to the pipeline:
  OK     = 1
  Cancel = 2
  Abort  = 3
  Retry  = 4
  Ignore = 5
  Yes    = 6
  No     = 7

  If no button is clicked, the return value is -1.

.OUTPUTS
    Null   = -1
    OK     = 1
    Cancel = 2
    Abort  = 3
    Retry  = 4
    Ignore = 5
    Yes    = 6
    No     = 7

.EXAMPLE
PS C:\> New-Popup -Message "The update script has completed" -Title "Finished" -Time 5

This will display a popup message using the default OK button and default 
Information icon. The popup will automatically dismiss after 5 seconds.

.NOTES
Original Source: http://powershell.org/wp/2013/04/29/powershell-popup/
Original Source: The Lonely Administrator's https://jdhitsolutions.com/blog/powershell/2976/powershell-popup/
Last Updated: April 8, 2013
Version     : 1.0
Similar to: Bill Riedy's https://www.powershellgallery.com/packages/PoshFunctions/2.2.1.6/Content/Functions%5CNew-Popup.ps1
#>

#------------------------------------------------------------------
# Display a Popup Message
#------------------------------------------------------------------
Function New-Popup {
    [CmdletBinding(SupportsShouldProcess = $true)]
    Param (
        [Parameter(Position=0, Mandatory=$true, HelpMessage="Enter a message for the popup")]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [Parameter(Position=1, Mandatory=$true, HelpMessage="Enter a title for the popup")]
        [ValidateNotNullOrEmpty()]
        [string]$Title,

        [Parameter(Position=2, HelpMessage="How many seconds to display? Use 0 to require a button click.")]
        [ValidateScript({$_ -ge 0})]
        [int]$Time = 0,

        [Parameter(Position=3, HelpMessage="Enter a button group")]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("OK", "OKCancel", "AbortRetryIgnore", "YesNo", "YesNoCancel", "RetryCancel")]
        [string]$Buttons = "OK",

        [Parameter(Position=4, HelpMessage="Enter an icon set")]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Stop", "Question", "Exclamation", "Information")]
        [string]$Icon = "Information"
    )
    Process {
        if ($PSCmdlet.ShouldProcess("Display Popup Message")) {
            # Convert buttons to their integer equivalents
            Switch ($Buttons) {
                "OK"               {$ButtonValue = 0}
                "OKCancel"         {$ButtonValue = 1}
                "AbortRetryIgnore" {$ButtonValue = 2}
                "YesNo"            {$ButtonValue = 4}
                "YesNoCancel"      {$ButtonValue = 3}
                "RetryCancel"      {$ButtonValue = 5}
            }
        }
        
        # Set an integer value for icon type
        Switch ($Icon) {
            "Stop"        {$IconValue = 16}
            "Question"    {$IconValue = 32}
            "Exclamation" {$IconValue = 48}
            "Information" {$IconValue = 64}
        }
        
        # Create the COM Object
        Try {
            $Wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
            # Button and icon type values are added together to create an integer value
            $Wshell.Popup($Message, $Time, $Title, $ButtonValue + $IconValue)
        } Catch {
            Write-Warning "Failed to create Wscript.Shell COM object"
            Write-Warning $_.Exception.Message
        }
    }
}
