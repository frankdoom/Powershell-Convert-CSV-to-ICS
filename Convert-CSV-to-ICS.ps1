# fonte dati http://it.cathopedia.org/wiki/Calendario_dei_santi


$InputFilename = $PSScriptRoot + "\template.csv"
$InputFilename = $PSScriptRoot + "\2018.csv"
$OutputFilename = $PSScriptRoot + "\output.ics"

#$CSV=Import-Csv $InputFilename
$CSV=Import-Csv -Delimiter ";" $InputFilename

# get CSV headers
$Headers = $CSV | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'

"BEGIN:VCALENDAR
PRODID:
VERSION:2.0
CALSCALE:GREGORIAN
METHOD:PUBLISH
X-WR-CALNAME:TEST IMPORT
X-WR-TIMEZONE:Europe/Rome
BEGIN:VTIMEZONE
TZID:Europe/Rome
X-LIC-LOCATION:Europe/Rome
BEGIN:DAYLIGHT
TZOFFSETFROM:+0100
TZOFFSETTO:+0200
TZNAME:CEST
DTSTART:19700329T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:+0200
TZOFFSETTO:+0100
TZNAME:CET
DTSTART:19701025T030000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
END:STANDARD
END:VTIMEZONE" > $OutputFilename

ForEach ($Line in $CSV) {
    "BEGIN:VEVENT" >> $OutputFilename
    ForEach ($FieldName in $Headers) {
        $($($FieldName)+":"+$($Line.$FieldName)) >> $OutputFilename
    }
    "TRANSP:TRANSPARENT" >> $OutputFilename
    "END:VEVENT" >> $OutputFilename
}

"END:VCALENDAR" >> $OutputFilename


# For ($i=0; $i -le $CSV.Count; $i++) {
#    echo $i
#    echo $CSV[$i]
#}