<#
.SYNOPSIS
    Converts a CSV file to an ICS calendar file.

.DESCRIPTION
    This script reads a CSV file containing event details and converts it to an ICS calendar file.

.PARAMETER InputFileBasename
    Name of the input CSV file.

.PARAMETER OutputFileBasename
    Name of the output ICS file. If not provided, the input name will be used.

.EXAMPLE
    .\Convert-CSV-to-ICS.ps1 -InputFileBasename 'events' -OutputFileBasename 'events'

.NOTES
    Author: Francesco Menin
    Date: 2025-03-21
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true, HelpMessage="Name of the input CSV file without extension.")]
    [string]$InputFileBasename,
    [Parameter(Mandatory=$false, HelpMessage="Name of the output ICS file without extension. If not provided, the input name will be used.")]
    [string]$OutputFileBasename = $null
)
#$Filename = "calendar_appointments"

# Se l'output filename non è fornito, usa lo stesso nome dell'input filename
if (-not $OutputFileBasename) {
    $OutputFileBasename = $InputFileBasename
}

$Delimiter = ','
$InputFilename = $PSScriptRoot + "\" + $InputFileBasename + ".csv"
$OutputFilename = $PSScriptRoot + "\" + $OutputFileBasename + ".ics"

#$CSV=Import-Csv $InputFilename
#$CSV=Import-Csv -Delimiter ";" -Path $InputFilename
$CSV=Import-Csv -Delimiter $Delimiter -Path $InputFilename

# get CSV headers
$Headers = $CSV | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'

"BEGIN:VCALENDAR
PRODID:
VERSION:2.0" > $OutputFilename

ForEach ($Line in $CSV) {
    $startDateTime = [datetime]::ParseExact("$($Line.'Start Date') $($Line.'Start Time')", 'dd/MM/yyyy HH:mm', $null)
    $endDateTime = [datetime]::ParseExact("$($Line.'End Date') $($Line.'End Time')", 'dd/MM/yyyy HH:mm', $null)

    "BEGIN:VEVENT" >> $OutputFilename
    "SUMMARY:$($(Line.Subject)" >> $OutputFilename
    "DTSTART:$($startDateTime.ToString('yyyyMMddTHHmmssZ'))" >> $OutputFilename
    "nDTEND:$($endDateTime.ToString('yyyyMMddTHHmmssZ'))" >> $OutputFilename
    "LOCATION:$($entry.Location)" >> $OutputFilename
    "DESCRIPTION:$($entry.Description)" >> $OutputFilename
    "END:VEVENT" >> $OutputFilename >> $OutputFilename
}

"END:VCALENDAR" >> $OutputFilename
