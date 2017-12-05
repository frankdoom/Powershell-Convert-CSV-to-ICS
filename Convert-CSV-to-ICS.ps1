$InputFilename = $PSScriptRoot + "\template.csv"
$OutputFilename = $PSScriptRoot + "\output.ics"
echo $OutputFilename
$CSV=Import-Csv $InputFilename
echo $CSV[1]

