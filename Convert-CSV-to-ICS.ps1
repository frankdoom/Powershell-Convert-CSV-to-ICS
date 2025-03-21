$Filename = "calendar_appointments"
$Delimiter = ','
$InputFilename = $PSScriptRoot + "\" + $Filename + ".csv"
$OutputFilename = $PSScriptRoot + "\" + $Filename + ".ics"

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
