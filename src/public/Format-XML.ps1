<#
.SYNOPSIS
    Formats an XML string with indentation and optional attribute formatting and XML declaration.

.DESCRIPTION
    The Format-XML function formats an XML string with a specified indentation level. It provides options to format attributes with new lines and to include or omit the XML declaration.

.PARAMETER xml
    The XML string to be formatted.

.PARAMETER indent
    The number of spaces to use for indentation. Default is 4.

.PARAMETER FormatAttributes
    A switch to format attributes with new lines.

.PARAMETER IncludeXmlDeclaration
    A switch to include the XML declaration.

.EXAMPLE
    Format-XML -xml $inputXML -indent 4

    Formats the XML string $inputXML with an indentation level of 4 spaces.

.EXAMPLE
    Format-XML -xml $inputXML -indent 4 -FormatAttributes

    Formats the XML string $inputXML with an indentation level of 4 spaces and formats attributes with new lines.

.EXAMPLE
    Format-XML -xml $inputXML -indent 4 -IncludeXmlDeclaration

    Formats the XML string $inputXML with an indentation level of 4 spaces and includes the XML declaration.

.NOTES
    Author: Brooks Vaughn
    Date: 2025-02-18 14:22:52
#>
function Format-XML {
    [CmdletBinding()]
    Param (
        [parameter(Mandatory=$true,ValueFromPipeLine=$true)]
        [xml]$xml,
        [parameter(Mandatory=$false,ValueFromPipeLine=$false)]
        $indent = 4,
        [switch]$FormatAttributes,
        [switch]$IncludeXmlDeclaration
    )
    Process {
        $StringWriter = New-Object System.IO.StringWriter
        $XmlWriterSettings = New-Object System.Xml.XmlWriterSettings
        $XmlWriterSettings.Indent = $true
        $XmlWriterSettings.IndentChars = " " * $indent
        If (-not [Sting]::IsNullOrEmpty($IncludeXmlDeclaration)) {
            $XmlWriterSettings.OmitXmlDeclaration = $true
        }
        If (-not [Sting]::IsNullOrEmpty($FormatAttributes)) {
            $XmlWriterSettings.NewLineOnAttributes = $true
        }
        $XmlWriter = [System.Xml.XmlWriter]::Create($StringWriter, $XmlWriterSettings)
        $xml.WriteTo($XmlWriter)
        $XmlWriter.Flush()
        $StringWriter.Flush()
        $XmlWriter.Close()
        $Result = $StringWriter.ToString()
        $StringWriter.Dispose()
        $Result
    }
}
