function Benchmark-Command {
<# HEADER
/*=====================================================================
Program Name            : Benchmark-Command.psm1
Purpose                 : Function to benchmark a command
Powershell Version:     : v2.0
Input Data              : N/A
Output Data             : N/A

Originally Written by   : Scott Bass
Date                    : 30OCT2014
Program Version #       : 1.0

=======================================================================

Modification History    :

=====================================================================*/

/*---------------------------------------------------------------------

=======================================================================

NOTE:

This script was modified from:
http://zduck.com/2013/benchmarking-with-Powershell/

---------------------------------------------------------------------*/
#>

<#
.SYNOPSIS
  Runs the given script block and returns the execution duration.
  Hat tip to StackOverflow. http://stackoverflow.com/questions/3513650/timing-a-commands-execution-in-powershell

.EXAMPLE
  Benchmark-Command { ping localhost -n 3 }
#>

#region Parameters
[CmdletBinding()]
param(
   [ScriptBlock]$Expression
   ,
   [int]$Samples = 1
   ,
   [Switch]$Silent
   ,
   [Switch]$Long
)
#endregion

#region Main
$timings = @()
do {
   $sw = New-Object Diagnostics.Stopwatch
   if ($Silent) {
      $sw.Start()
      $null = & $Expression
      $sw.Stop()
      Write-Host "." -NoNewLine
   }
   else {
      $sw.Start()
      & $Expression
      $sw.Stop()
   }
   $timings += $sw.Elapsed

   $Samples--
}
while ($Samples -gt 0)

Write-Host

$stats = $timings | Measure-Object -Average -Minimum -Maximum -Property Ticks

# Print the full timespan if the $Long switch was given.
if ($Long) {
   "Avg: {0:dd\.hh\:mm\:ss\.ff}" -f $(New-Object System.TimeSpan $stats.Average)
   "Min: {0:dd\.hh\:mm\:ss\.ff}" -f $(New-Object System.TimeSpan $stats.Minimum)
   "Max: {0:dd\.hh\:mm\:ss\.ff}" -f $(New-Object System.TimeSpan $stats.Maximum)
}
else {
   # Otherwise just print the milliseconds which is easier to read.
   #"Avg: {0:N2} ms" -f $((New-Object System.TimeSpan $stats.Average).TotalMilliseconds)
   #"Min: {0:N2} ms" -f $((New-Object System.TimeSpan $stats.Minimum).TotalMilliseconds)
   #"Max: {0:N2} ms" -f $((New-Object System.TimeSpan $stats.Maximum).TotalMilliseconds)

   # Otherwise just print the seconds which is easier to read.
   "Avg: {0:N2} secs" -f $((New-Object System.TimeSpan $stats.Average).TotalSeconds)
   "Min: {0:N2} secs" -f $((New-Object System.TimeSpan $stats.Minimum).TotalSeconds)
   "Max: {0:N2} secs" -f $((New-Object System.TimeSpan $stats.Maximum).TotalSeconds)
}
#endregion
}
