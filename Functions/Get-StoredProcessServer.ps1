# Function to get the SAS Stored Process Server
function Get-StoredProcessServer {
   param (
      [ValidateSet("Lev1","Lev2","Lev9")]
      [String]$lev="Lev1"
      ,
      [String[]]$vars="Name,ProcessId,commandLine,creationDate"
   )
# I know I could put all this in a pipeline, but this seemed easier for others to understand and maintain
$sasprocesses=Get-WMIObject win32_process -filter "name='sas.exe'" 
$storedprocesses=$sasprocesses | where {$_.commandline -like "*StoredProcessServer*"}
$storedprocess=$storedprocesses | where {$_.commandline -like "*\$lev\*"}
$storedprocess | format-list ($vars -split ",")

[datetime]::ParseExact(($storedprocess.CreationDate.Split("."))[0],”yyyyMMddhhmmss”,$null)
}
