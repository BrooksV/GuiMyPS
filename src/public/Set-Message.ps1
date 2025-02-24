<#
.SYNOPSIS
Sets a message in a WPF String control using DispatcherTimer to Set Message Box via WPF String control.

.DESCRIPTION
This function sets a message in a WPF String control and allows for various options such as clearing the message, adding new lines before or after, and avoiding new lines.

Requires Defining an XAML: Element named Messages which gets updated by a DispatchTimer
    <Window.Resources>
        <system:String x:Key="Messages">Messages:</system:String>
    <Window.Resources>

    <ScrollViewer x:Name="_svMessages" Padding="0,0,0,0" ScrollViewer.VerticalScrollBarVisibility="Visible" ScrollViewer.HorizontalScrollBarVisibility="Disabled">
        <TextBox x:Name="_tbxMessages" Margin="5,2,5,5" Padding="0,0,0,0" 
            Text="{DynamicResource ResourceKey=Messages}"
            ScrollViewer.CanContentScroll="False"
            HorizontalAlignment="Stretch" VerticalAlignment="Stretch" 
            VerticalContentAlignment="Top"
            TextWrapping="WrapWithOverflow" AcceptsReturn="True" Cursor="Arrow"
        />
    </ScrollViewer>    

.PARAMETER Message
The message to set in the WPF String control.

.PARAMETER Clear
Clears the current message.

.PARAMETER NoNewline
Specifies not to add a newline after the message.

.PARAMETER NewlineBefore
Adds a newline before the message.

.PARAMETER NewlineAfter
Adds a newline after the message.

.EXAMPLE
Set-Message -Resources $syncHash.Form.Resources -Message 'This is a test message'
Set-Message -Resources $syncHash.Form.Resources -Message 'This is a test message' -Clear

.NOTES
Author: Brooks Vaughn
Date: "22 February 2025"
#>

Function Set-Message {
    [CmdletBinding(SupportsShouldProcess = $true)]
    Param (
        [Object]$Resources = $SyncHash.Form.Resources,
        [string]$Message = [string]::Empty,
        [switch]$Clear,
        [switch]$NoNewline,
        [switch]$NewlineBefore,
        [switch]$NewlineAfter
    )

    If ($PSCmdlet.ShouldProcess("Setting message")) {
        If ($Clear) {
            $Message = [string]::Empty
            $Resources["Messages"] = [string]::Empty
        }
    }

    If ($NewlineBefore) {
        $Resources["Messages"] += [System.Environment]::NewLine
    }

    If (-not [string]::IsNullOrEmpty($Message)) {
        If ($NoNewline) {
            $Resources["Messages"] += $Message
        } Else {
            $Resources["Messages"] += $Message + [System.Environment]::NewLine
        }
        Write-Host $Message -NoNewline:$NoNewline
    }

    If ($NewlineAfter) {
        $Resources["Messages"] += [System.Environment]::NewLine
    }
}
