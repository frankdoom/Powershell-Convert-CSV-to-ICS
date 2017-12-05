$InputFilename = $PSScriptRoot + "\template.csv"
$OutputFilename = $PSScriptRoot + "\output.ics"
echo $OutputFilename
$CSV=Import-Csv $InputFilename

# get CSV headers
$Headers = $CSV | Get-Member -MemberType NoteProperty


For ($i=0; $i -le $CSV.Count; $i++) {
    echo $i
    echo $CSV[$i]
}