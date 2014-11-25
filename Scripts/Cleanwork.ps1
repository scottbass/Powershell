<# HEADER
/*=====================================================================
Program Name            : cleanwork.ps1
Purpose                 : Deletes orphaned SAS Work and SAS Utility
                          directories

Originally Written by   : Scott Bass
Date                    : 18SEP2012
Program Version #       : 1.0

=======================================================================

Modification History    : Original version

=====================================================================*/
#>

<#
.SYNOPSIS
Deletes orphaned SAS Work and/or SAS Utility directories and their contents.
.DESCRIPTION
Deletes orphaned SAS Work and/or SAS Utility directories and their contents.
"Orphaned" means those directories that do not have a matching SAS process actively running.
.NOTES
See http://support.sas.com/documentation/cdl/en/hostwin/63285/HTML/default/viewer.htm#a003182430.htm
for details on the naming convention for SAS Work and Utility directories.

This script does not delete the temporary SAS files referenced in the
above documentation.  It only deletes orphaned SAS Work and SAS Utility
directories.

This script only runs on Windows, and requires Powershell 2.0 or above to be installed.
.PARAMETER Dirs
Specifies the root directories to search for SAS Work or Utility sub-directories.
.PARAMETER Details
Prints diagnostic information in the Powershell console window.
.PARAMETER WhatIf
Explains what will happen if the command is executed, without actually executing the command.
.PARAMETER Confirm
Requests confirmation before the operation is executed.
.PARAMETER Verbose
Generates detailed information about the operation, much like tracing or a transaction log.
.INPUTS
Directory path(s) to specified directory(ies), which contain sub-directories whose names match _TD* or SAS_util*.
.OUTPUTS
None
.EXAMPLE
.\cleanwork.ps1
Prompts for the root directory(ies), then deletes orphaned SAS Work and SAS Utility directories.
During prompting, enter a blank value (just hit enter) to execute the script,
or press Cntl-C to abort processing.
.EXAMPLE
.\cleanwork.ps1 C:\Temp\SASWork -or- .\cleanwork.ps1 -dirs C:\Temp\SASWork
Delete all orphaned SAS Work and SAS Utility directories under C:\Temp\SASWork.
.EXAMPLE
.\cleanwork.ps1 C:\Temp\SASWork, C:\Temp\SASUtil -or- .\cleanwork.ps1 -dirs C:\Temp\SASWork, C:\Temp\SASUtil
Delete all orphaned SAS Work and SAS Utility directories under C:\Temp\SASWork and C:\Temp\SASUtil.
Use a comma to separate directories.
.EXAMPLE
.\cleanwork.ps1 C:\ -or- .\cleanwork.ps1 -dirs C:\
Delete all orphaned SAS Work and SAS Utility directories on the entire C: drive.
.EXAMPLE
.\cleanwork.ps1 C:\Temp\SASWork -whatif
Display what would happen if this script were executed, without actually deleting any directories.
.EXAMPLE
.\cleanwork.ps1 C:\Temp\SASWork -confirm
Requests confirmation for each directory targetted for deletion.
.EXAMPLE
.\cleanwork.ps1 C:\Temp\SASWork -verbose
Display additional tracing information as the directories are deleted.
#>

#region Parameters
[CmdletBinding(
   SupportsShouldProcess=$true,
   ConfirmImpact="Medium"
)]
Param(
   [Parameter(
      Position=0,
      Mandatory=$true,
      ValueFromPipeline=$false,
      ValueFromPipelineByPropertyName=$false
   )]
   [String[]] $dirs
   ,
   [Switch] $details
)
#endregion

#region Main
# initialize variables
$allsasdirs    = @() # all SASWork and SASUtil directories found
$matchsasdirs  = @() # SASWork and SASUtil directories with matching SAS process IDs
$delsasdirs    = @() # SASWork and SASUtil directories to be deleted (the diff between $allsasdirs and $matchsasdirs)
$pids          = @() # SAS process IDs

# Get a list of all SASWork and SASUtil directories.
# Recursively search the list of directories specified on the command line,
# only keeping those directories whose names begin with _TD or SAS_util (case-insensitive)
foreach ($dir in $dirs) {
   Get-ChildItem $dir -Recurse -ErrorAction SilentlyContinue | `
      Where {$_.psIsContainer} | `
      Where {$_.Name -match "^_TD|^SAS_util"} | `
      ForEach-Object {$allsasdirs+=$_}
}

# Get a list of SAS Process IDs
$pids=Get-Process -Name sas -ErrorAction SilentlyContinue

# Find directories with matching SAS process IDs.
foreach ($id in $pids) {
   foreach ($dir in $allsasdirs) {
      # SASWork: example: _TD12345_HOSTNAME_
      if ($dir.Name -match "^(_TD)(\d+)_(\w+)$") {
         $_dirpid=$Matches[2]
         $_pid=$id.id
         if ($_dirpid -eq $_pid) {$matchsasdirs+=$dir}
      }
      # SASUtil: example: SAS_util00010000088C_HOSTNAME
      # SASUtil: example: SAS_util<serial><PID: hex8.>_<hostname>
      if ($dir.Name -match "^(SAS_util)(.*?)_(\w+)$") {
         $_dirpid=$Matches[2]
         $_dirpid=$_dirpid.substring($_dirpid.length-8,8) # get the last 8 characters from the right
         $_pid=$id.id
         $_pidhex="{0:X8}" -f [Int]$_pid
         if ($_dirpid -eq $_pidhex) {$matchsasdirs+=$dir}
      }
   }
}

# Get a list of SASWork and SASUtil directories to be deleted.
# This will be those directories in $allsasdirs without a match in $matchsasdirs
$compare=Compare-Object $matchsasdirs $allsasdirs
$compare | ForEach-Object {if ($_.SideIndicator -eq '=>') {$delsasdirs+=$_.InputObject}}

# print debugging information if -details switch was specified
if ($PSBoundParameters.Details) {
   $fmt=@{Expression={$_.ID};Label="SAS Process ID";width=20},@{Expression={"{0:X8}" -f $_.ID};Label="SAS Process ID (hex)";width=20}
   $line="================================================================="
   "SAS Work/Utility Directories Found:"
   $allsasdirs | Format-Table -Property Fullname -AutoSize
   $line
   "SAS Process IDs:"
   $pids | Format-Table $fmt -AutoSize
   $line
   "SAS Work/Utility Directories with Active SAS Process:"
   $matchsasdirs | Format-Table -Property Fullname -AutoSize
   $line
   "SAS Work/Utility Directories to be deleted:"
   $delsasdirs | Format-Table -Property Fullname -AutoSize
}

# delete the orphaned directories
foreach ($dir in $delsasdirs) {
   if ($PSCmdlet.ShouldProcess(
      "$dir","Remove-Item"))
   {
      try {
         If (Test-Path $dir.Fullname -pathtype Container) {
            Get-ChildItem $dir.Fullname -recurse | Get-Acl | Select-Object @{Label="Path";Expression={Convert-Path $_.Path}},Owner
            Write-Output "`n"
         }
         If (Test-Path $dir.Fullname -pathtype Container) {
            Remove-Item $dir.Fullname `
               -Recurse `
               -Verbose:($PSBoundParameters.Verbose -eq $true) `
               -ErrorAction "Continue"
         }
      }
      catch {
         Write-Error "$_"
      }
   }
}
#endregion
