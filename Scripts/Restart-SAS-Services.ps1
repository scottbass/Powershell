<#
Note:  The brackets [ ] in the SAS service names are regular expression metacharacters.
They cannot be referenced directly in the service names.
Use the question mark wildcard instead.
#>

$ErrorActionPreference  = "Continue"
$DebugPreference        = 2
$VerbosePreference      = 2
$WarningPreference      = 2

# === Set Sleep Length (in Seconds) ===
$Sleep                  = 60

# === Increase Batch Mode Console Line Size ===
$pshost              = get-host
$pswindow            = $pshost.ui.rawui
$newsize             = $pswindow.buffersize
$newsize.height      = 3000
$newsize.width       = 255
$pswindow.buffersize = $newsize

# === Create log file ===
$log=(Get-ChildItem $MyInvocation.MyCommand.Definition).FullName
$log=$log + ".log"
# Start-Transcript "$log"

# === Print Date and Time ===
(Get-Date).DateTime
Write-Output "`n"

# === Stop Dependent Services ===
Stop-Service -name "SAS*Lev1*Connect Spawner*"
Stop-Service -name "SAS*Lev2*Connect Spawner*"
Stop-Service -name "SAS*Lev9*Connect Spawner*"

Stop-Service -name "SAS*Lev1*Object Spawner*"
Stop-Service -name "SAS*Lev2*Object Spawner*"
Stop-Service -name "SAS*Lev9*Object Spawner*"

# === Stop Metadata Server ===
Stop-Service -name "SAS*Lev1*Metadata Server*" -Force
Stop-Service -name "SAS*Lev2*Metadata Server*" -Force
Stop-Service -name "SAS*Lev9*Metadata Server*" -Force

# === Sleep for X Seconds ====
Start-Sleep -Seconds $Sleep

# === List status of SAS Services (all should be stopped) ====
Write-Output "----------------+"
Write-Output "Stopped Services:"
Write-Output "----------------+"
Get-Service -name "SAS*Connect*", "SAS*Object*", "SAS*Metadata*" | Format-Table -AutoSize

# === Start Metadata Server ===
Start-Service -name "SAS*Lev1*Metadata Server*"
Start-Service -name "SAS*Lev2*Metadata Server*"
Start-Service -name "SAS*Lev9*Metadata Server*"

# === Start Dependent Services ===
Start-Service -name "SAS*Lev1*Object Spawner*"
Start-Service -name "SAS*Lev2*Object Spawner*"
Start-Service -name "SAS*Lev9*Object Spawner*"

Start-Service -name "SAS*Lev1*Connect Spawner*"
Start-Service -name "SAS*Lev2*Connect Spawner*"
Start-Service -name "SAS*Lev9*Connect Spawner*"

# === Sleep for X Seconds ====
Start-Sleep -Seconds $Sleep

# === List status of SAS Services (all should be started) ====
Write-Output "----------------+"
Write-Output "Started Services:"
Write-Output "----------------+"
Get-Service -name "SAS*Connect*", "SAS*Object*", "SAS*Metadata*" | Format-Table -AutoSize

# === Test that all services are running ===
$stopped=Get-Service -name "SAS*Connect*", "SAS*Object*", "SAS*Metadata*" | Where {$_.Status -ne "Running"}
if ($stopped -ne $null) {
   Write-Output "ERROR: The below SAS Services are not running:`n"
   $stopped | Format-Table -AutoSize
}

# === Print Date and Time ===
(Get-Date).DateTime

# === If some services are not running throw an error ===
if ($stopped -ne $null) {
   Throw "One or more SAS services were not restarted."
}

# === Close log file ===
# Stop-Transcript
