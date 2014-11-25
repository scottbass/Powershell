function Run-Robocopy {
<# HEADER
/*=====================================================================
Program Name            : Run-Robocopy.ps1
Purpose                 : Run Robocopy with standard parameters
Powershell Version:     : v2.0
Input Data              : N/A
Output Data             : N/A

Originally Written by   : Scott Bass
Date                    : 29OCT2014
Program Version #       : 1.0

=======================================================================

Modification History    :

=====================================================================*/

/*---------------------------------------------------------------------

This script runs Robocopy with standard parameters that are best for
logging and performance.

Note that by default Robocopy will not replace already existing files
in the destination folder that have the same filesize and modification
time of the source file.  What this means is the first copy/backup may
take a long time, but subsequent copies/backups will likely go much
quicker, depending on the data that has actually changed in the source.
=======================================================================

NOTE:

http://www.sevenforums.com/tutorials/187346-robocopy-create-backup-script.html
http://ss64.com/nt/robocopy.html

Robocopy parameters (subset of the parameters we're interested in):

robocopy <source_folder> <destination_folder> <wildcard> /COPY:DAT /DCOPY:T /E /MIR /NP /ZB /TEE /MT:8 /LOG:backup_log.txt

<source_folder>:      The source folder for the files, eg. \\aumelbcasas01\e$
<destination_folder>: The destination folder for the files, eg. \\melmp0505\e$
                      or just E:\ if running the script on the target machine.
<wildcard>:           Wildcard for the files to be copied (*.* by default)

/COPY:DAT             What to COPY (default is /COPY:DAT)
                      (copyflags : D=Data, A=Attributes, T=Timestamps
                       S=Security=NTFS ACLs, O=Owner info, U=aUditing info).

/DCOPY:T              Copy Directory Timestamps.

/E                    Copy Subfolders, including Empty Subfolders.

/MIR                  MIRror a directory tree - equivalent to /PURGE plus all subfolders (/E)
                      (PURGE=Delete dest files/folders that no longer exist in source.)

/NP                   No Progress - don’t display % copied.

/ZB                   Use restartable mode; if access denied use Backup mode.

/TEE                  Output to console window, as well as the log file.

/MT:#                 Multithreaded copying, n = no. of threads to use (1-128)
                      default = 8 threads, not compatible with /IPG and /EFSRAW
                      The use of /LOG is recommended for better performance.

/XF                   eXclude Files matching given names/paths/wildcards.

/XD                   eXclude Directories matching given names/paths.
                      XF and XD can be used in combination  e.g.
                      ROBOCOPY c:\source d:\dest /XF *.doc *.xls /XD c:\unwanted /S

/R:#                  Number of Retries on failed copies - default is 1 million.

/W:#                  Wait time between retries - default is 30 seconds.

/NFL                  No File List - don’t log file names.

/NDL                  No Directory List - don’t log directory names.

/LOG:                 Output status to LOG file (overwrite existing log).

/LOG+:                Output status to LOG file (append to existing log).

/ETA                  Show Estimated Time of Arrival of copied files.

---------------------------------------------------------------------*/
#>

<#
.SYNOPSIS
Run Robocopy Function

.DESCRIPTION
Run Robocopy with options optimized for logging and performance.

.PARAMETER  Source
Source folder to copy files from

.PARAMETER  Destination
Destination folder to copy files to

.PARAMETER  Wildcard
Filename wildcard (default *.*)

.PARAMETER  Options
Robocopy options (default /TEE /COPY:DAT /DCOPY:T /MIR /NP /ZB /R:0 /W:0 /ETA)

.PARAMETER  Log
Full path to Robocopy log file (default backup.log in the current directory)
Do not write the log in the destination folder.
In that scenario, if /MIR is specified (it is by default),
you will have a deadlock condition and the processing will hang.

.PARAMETER  Useropts
User options to augment or replace the default options

.EXAMPLE
Run-Robocopy \\aumelbcasas01\e$ E:\ R:\Temp\Run-Robocopy_E.log
Alternatives:
Run-Robocopy \\aumelbcasas01\e$ E:\ -log R:\Temp\Run-Robocopy_E.log
Run-Robocopy -source \\aumelbcasas01\e$ -dest E:\ -log R:\Temp\Run-Robocopy_E.log

Description
-----------
Copy all files from \\aumelbcasas01\e$ to the E:\ drive on the local machine.

.EXAMPLE
Run-Robocopy \\aumelbcasas01\e$ \\melmp0505\e$ *.sas R:\Temp\Run-Robocopy_E.log
Alternatives:
Run-Robocopy -source \\aumelbcasas01\e$ -dest \\melmp0505\e$ -wildcard *.sas -log R:\Temp\Run-Robocopy_E.log

Description
-----------
Copy only SAS program files from \\aumelbcasas01\e$ to \\melmp0505\e$

.EXAMPLE
Run-Robocopy -source \\aumelbcasas01\e$ -dest E:\ -log R:\Temp\backup.log -useropts /QUIT
Alternatives:
Run-Robocopy -source \\aumelbcasas01\e$ -dest E:\ -log R:\Temp\backup.log -useropts "/QUIT"

Description
-----------
Copy all files from \\aumelbcasas01\e$ to the E:\ drive on the local machine,
augmenting the default options with the user specified options.

.EXAMPLE
Run-Robocopy -source \\aumelbcasas01\e$ -dest E:\ -log R:\Temp\backup.log -options "" -useropts /COPYALL /E /NP /R:3 /W:10 /QUIT
Alternatives:
Run-Robocopy -source \\aumelbcasas01\e$ -dest E:\ -log R:\Temp\backup.log -options "" -useropts "/COPYALL /E /NP /R:3 /W:10 /QUIT"

Description
-----------
Copy all files from \\aumelbcasas01\e$ to the E:\ drive on the local machine,
replacing all the default options with the user specified options.

.EXAMPLE
$datestamp = Get-Date -f "yyyyMMdd"; $datetimestamp = Get-Date -f "yyyyMMdd_HHmmss"
Run-Robocopy -source \\aumelbcasas01\e$ -dest E:\ -log R:\Temp\backup_${datestamp}.log
Run-Robocopy -source \\aumelbcasas01\e$ -dest E:\ -log R:\Temp\backup_${datetimestamp}.log

Run-Robocopy -source \\aumelbcasas01\e$ -dest E:\ -log R:\Temp\backup_$(Get-Date -f "yyyyMMdd").log
Run-Robocopy -source \\aumelbcasas01\e$ -dest E:\ -log R:\Temp\backup_$(Get-Date -f "yyyyMMdd_HHmmss").log

Description
-----------
Use this approach to time stamp the Robocopy log file.

.EXAMPLE
Run-Robocopy \\aumelbcasas01\e$ E:\ R:\Temp\Run-Robocopy_E.log -whatif

Description
-----------
Display what Run-Robocopy would do, without actually executing the command.

.EXAMPLE
Run-Robocopy \\aumelbcasas01\e$ E:\ R:\Temp\Run-Robocopy_E.log -confirm

Description
-----------
Request confirmation before executing the command.

.NOTES
Note that by default Robocopy will not replace already existing files
in the destination folder that have the same filesize and modification
time of the source file.  What this means is the first copy/backup may
take a long time, but subsequent copies/backups will likely go much
quicker, depending on the data that has actually changed in the source.
#>

#region Parameters
[CmdletBinding(SupportsShouldProcess=$true)]
param(
   [Parameter(
      Mandatory=$true
   )]
   [String]$source
   ,
   [Parameter(
      Mandatory=$true
   )]
   [String]$destination
   ,
   [Parameter(
      Mandatory=$false
   )]
   [String]$wildcard
   ,
   [Parameter(
      Mandatory=$false
   )]
   [String[]]$options='/TEE /COPY:DAT /DCOPY:T /MIR /NP /ZB /R:0 /W:0 /ETA /XD "`$RECYCLE.BIN" "RECYCLER" "System Volume Information" '
   ,
   [Parameter(
      Mandatory=$false
   )]
   [String]$log="backup.log"
   ,
   [Parameter(ValueFromRemainingArguments=$true)]
   [String[]]$useropts
)
#endregion

#region Main
$cmd="robocopy ""$source"" ""$destination"" $wildcard $options $useropts"
if ($log) {$cmd += " /LOG:$log"}

if ($PSCmdlet.ShouldProcess(
   "`n$cmd"
))
{
   Invoke-Expression $cmd
}
#endregion
}
