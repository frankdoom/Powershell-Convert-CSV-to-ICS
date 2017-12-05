$InputFilename = $PSScriptRoot + "\template.csv"
$OutputFilename = $PSScriptRoot + "\output.ics"

$CSV=Import-Csv $InputFilename

# get CSV headers
$Headers = $CSV | Get-Member -MemberType NoteProperty

ForEach ($Line in $CSV) {
    echo $Line
}


# For ($i=0; $i -le $CSV.Count; $i++) {
#    echo $i
#    echo $CSV[$i]
#}