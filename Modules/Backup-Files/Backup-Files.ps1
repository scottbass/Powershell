<# HEADER
/*=====================================================================
Program Name            : Backup-Files.ps1
Purpose                 : Function to archive files by various criteria.
Powershell Version:     : v2.0
Input Data              : (Optional) List of files in the pipeline
Output Data             : Archive of selected files.

Originally Written by   : Scott Bass
Date                    : 01MAR2013
Program Version #       : 1.0

=======================================================================

Modification History    :

=====================================================================*/
#>

<#
.SYNOPSIS
Script to archive files by various criteria.

.DESCRIPTION
Backup-Files is used to archive files matching various selection criteria.

The file selection criteria is implemented by the Select-Files function.
This function must be loaded in order for this script to work.

The selection criteria is documented in the Select-Files function.
Type "help Select-Files" for more information.
Note that this script does not support the Select-Files -type parameter.

If the archive process is successful, the files that were processed are returned
to the pipeline.  Thus, the results can be assigned to a variable or piped to
other cmdlets or scripts.  For example:  Backup-Files | Remove-Files.

This script uses the 7-Zip program to archive the files.  The script code would
need to be changed if another archive utility is used.

.PARAMETER  Path
Specifies a path to one or more locations. Wildcards are permitted.
The default location is the current directory (.).

.PARAMETER  Extension
Lists only files that have the specified extension.  Multiple extensions are
allowed, separated by commas.  Specify ONLY the extension, including the leading
period, for example .txt, not *.txt.

.PARAMETER  Filter
Specifies a filter in the provider's format or language. The value of this
parameter qualifies the Path parameter. The syntax of the filter, including the
use of wildcards, depends on the provider.

Filters are more efficient than other parameters, because the provider applies
them when retrieving the objects, rather than having Windows PowerShell filter
the objects after they are retrieved.

NOTE: The Filter parameter will perform faster than the equivalent Include or
Exclude parameter.  However, only one filter can be specified for the Filter
parameter, whereas multiple filters can be specified for the Include or Exclude
parameters.  See examples for more details.

.PARAMETER  Include
Retrieves only the specified items. The value of this parameter qualifies the
Path parameter. Enter a path element or pattern, such as "*.txt". Wildcards are
permitted.  Multiple values are permitted, separated by commas.

NOTE: The Include parameter is effective only when the command includes the
Recurse parameter or the path leads to the contents of a directory, such as
C:\Windows\*, where the wildcard character specifies the contents of the
C:\Windows directory.

.PARAMETER  Exclude
Omits the specified items. The value of this parameter qualifies the Path
parameter. Enter a path element or pattern, such as "*.txt". Wildcards are
permitted.  Multiple values are permitted, separated by commas.

NOTE: Unlike the Include parameter, the Exclude parameter is effective when the
path specifies a directory path only, i.e. NOT the contents of a directory.  If
the path specifies the contents of a directory, the Exclude parameter will also
recurse sub-directories, which is usually not what you want.

.PARAMETER  Name
Retrieves only the names of the items in the locations. If you pipe the output
of this command to another command, only the item names are sent.

.PARAMETER  Recurse
Gets the items in the specified locations and in all child items of the
locations. Recurse works only when the path points to a container that has child
items, such as C:\Windows or C:\Windows\*, and not when it points to items that
do not have child items, such as C:\Windows\*.exe.

.PARAMETER  Older
Lists only files older than the specified date (exclusive).
Specify the desired date in DDMONYY(YY) or DD/MM/YY(YY) format,
for example 25dec12, 25DEC2012, 25/12/12, or 25/12/2012.

.PARAMETER  Newer
Lists only files newer than the specified date (exclusive).
Specify the desired date in DDMONYY(YY) or DD/MM/YY(YY) format,
for example 25dec12, 25DEC2012, 25/12/12, or 25/12/2012.

.PARAMETER  Between
Lists only files between two dates (inclusive).
Specify the start and end dates separated by a comma,
for example 01jan12,01feb12, or "01jan2012 00:00:00","31jan2012 23:59:59"

This parameter is a convenience to simulate the SQL between operator.
The same results could be derived by combining -newer and -older. In theory some
date fiddling might be required to return the same results, since -between is
inclusive of the date boundaries, while -newer and -older are not.  In practice,
since we most often filter by date, not datetime, and since file timestamps are
datetime, the level of granularity is such that both approaches usually return
the same results.

If -between is specified, do not specify -older or -newer, and vice versa.
However, no error checking is done to prohibit specifying conflicting
parameters, but you likely will get undesired (i.e. no) results.

.PARAMETER  Age
Lists only files that are "older" or "younger" than the specified number of
days, as measured from today.
A positive value lists files older than the specified number of days.
A negative value lists files younger than the specified number of days.

.PARAMETER  Size
Lists only files that are larger or smaller than the specified size.
A positive value lists those files larger than the specified size.
A negative value lists those files smaller than the specified size.

NOTE: Powershell mnemonics are allowed as size specifications, for example
1kb, 25MB, 3.5GB, 1.2Tb.

NOTE2: Due to a bug in Powershell, if specifying negative sizes, enclose the
value in parentheses, for example -size (-10kb) will return all files smaller
than 10kb.

.PARAMETER  Num
Lists only the specified number of files.
A positive value lists the X number of <oldest|largest> files, for example the 10 oldest files.
A negative value lists the X number of <youngest|smallest> files, for example the 10 smallest files.

The selection criteria is as follows:

If -size was specified, return the X number of <largest|smallest> files,
for example the 10 largest or 10 smallest files.

Otherwise, return the X number of <oldest|youngest> files,
for example the 10 oldest or 10 youngest files.

If you want the X largest files, specify -size 1, for example -size 1 -num 10.
If you want the X smallest files, specify -size 1, for example -size 1 -num -10.

.PARAMETER  WhichDate
Specifies which file object date property is used for those operations that are date based
(-older, -newer, -age, -num).
Valid values are "created", "accessed", or "modified".
The default value is "modified".

.PARAMETER  ScriptBlock
This parameter is for advanced users only!

Specifies an arbitary script block that can be used to filter the list of files returned.

Usually this would be used for advanced filtering based on date values, for example
LastWriteTime = Q1, or LastWriteTime = Sunday.  However, any queries against any
properties of a System.IO.FileInfo (or System.IO.DirectoryInfo) object can be specified.
Type gci . | Get-Member for more information.

This parameter is most useful when this script is embedded within another script,
to return a list of files on which to execute some action, and where advanced, dynamic
filtering is required.

The ScriptBlock parameter must be specified as a scriptblock, that is, within braces, for example
-ScriptBlock {Get-Date $_.LastWriteTime -gt Get-Date 01Jan2013}.

See http://ss64.com/ps/syntax-dateformats.html and http://ss64.com/bash/date.html#format
for date format and uformat pattern examples.

An internal hash table $QTR is created to facilitate filtering by quarter.
See script code for more details.

See examples in the help for Select-Files.

.PARAMETER Archive
Path to a new or existing archive file location.

.PARAMETER Log
Path to a log file which saves console messages generated by 7-Zip.

.PARAMETER Append
Switch parameter to append the 7-Zip console messages to the log file.
This parameter is ignored unless a log file is specified.

.PARAMETER WithPath
Switch parameter to include the directory path structure within the archive file.

.EXAMPLE
Backup-Files . C:\Temp\archive.7z
Backup-Files . -archive C:\Temp\archive.7z
Backup-Files -archive C:\Temp\archive.7z.

Description
-----------
Archive files in the current directory, creating the C:\Temp\archive.7z file.

.EXAMPLE
Backup-Files -path C:\Temp -filter *.txt -archive C:\Temp\archive.7z -log C:\Temp\archive.log -append

Description
-----------
Archive *.txt files in the C:\Temp directory, creating the C:\Temp\archive.7z file,
appending the console messages generated by 7-Zip to C:\Temp\archive.log.

.EXAMPLE
Backup-Files -path C:\Temp\Temp1,C:\Temp\Temp2 -exclude *.7z,*.log -archive C:\Temp\archive.7z -withpath -recurse

Description
-----------
Archive files in the C:\Temp\Temp1 and C:\Temp\Temp2 directories, including all
sub-directories, excluding *.7z and *.log files, creating the C:\Temp\archive.7z
fiile, and embedding the directory structure within the archive.

.EXAMPLE
Backup-Files -path C:\Temp -archive whatever -whatif
Backup-Files -path C:\Temp -archive whatever -whatif -withpath
Backup-Files -path C:\Temp -archive whatever -whatif -withpath -recurse

Description
-----------
Display what *would* happen (-whatif) in each of these examples.
Does not actually execute the archive process.

.EXAMPLE
Backup-Files -path C:\Temp -archive C:\Temp\archive.7z -confirm

Description
-----------
Lists the files that will be archived, giving you the opportunity to confirm the
operation (-confirm).  Note this is an "all or nothing" confirmation.  Review
all files listed before choosing "Yes".

#>

#region Parameters
[CmdletBinding()]
param(
   ### Get-ChildItem parameters
   [Alias("Fullname")]
   [Parameter(
      Position=0,
      ValueFromPipeline=$true,
      ValueFromPipelineByPropertyName=$true
   )]
   [Object[]]$Path
   ,
   [ValidatePattern("^\.[a-z]{2,5}")]
   [String[]]$Extension
   ,
   [String]$Filter
   ,
   [String[]]$Include
   ,
   [String[]]$Exclude
   ,
   [Switch]$Name
   ,
   [Switch]$Recurse
   ,

   ### Additional filtering parameters
   [DateTime]$Older
   ,
   [DateTime]$Newer
   ,
   [DateTime[]]$Between
   ,
   <#
   [ValidateSet("file","dir")]
   [String]$Type
   ,
   #>
   [Int]$Size
   ,
   [Int]$Age
   ,
   [Int]$Num
   ,
   [ValidateSet("created","accessed","modified")]
   [String]$WhichDate="modified"
   ,
   [ScriptBlock]$ScriptBlock
   ,

   # Parameters for Backup-Files
   [Parameter(Position=1, Mandatory=$true)]
   [String]$Archive
   ,
   [String]$Log
   ,
   [Switch]$Append
   ,
   [Switch]$WithPath
)
#endregion

#region Functions
#endregion

#region Main
begin {
   $ErrorActionPreference = "Stop"

   # Specify the path to your archive program.  Obviously we're using 7-zip.
   $ArchiveProgram = "D:\Program Files\7-Zip\7z.exe"

   # Specify any archive program parameters
   $ArchiveProgramParms = "a"

   # Array to hold files
   $files = @()
}

process {
   foreach ($file in $Path) {
      if ($file -eq $null)    {}
      else {
         switch -wildcard ($file.GetType().FullName) {
            "System.String"   {$files += $file}
            "System.IO.*"     {$files += $file.fullname}
            default           {throw "Invalid -Path: $objectType is invalid."}
         }
      }
   }
}

end {
   # If there are no files to process then abort.
   if (-not $files) {Write-Warning "There are no files to process."; return $null}

   # 7-zip doesn't like directories, only the directory paths to the files.
   # If directories are specified, 7-zip will recurse the directories, whether you want recursion or not.
   # Use the -recurse command line switch if you want to explicitly recurse the directory structure.

   # Now filter the files
   # Pass the script parameters to Select-Files by splatting the $PSBoundParameters collection.
   $PSBoundParameters.Item("Path")=$files
   $files = Select-Files @PSBoundParameters -Type file

   # If there are no files to process then abort.
   if (-not $files) {Write-Warning "There are no files to process."; return $null}

   # If -whatif was specified print what we would do and return
   # If -confirm was specified allow the user to select which files to process
   $temp = @()
   foreach ($file in $files) {
      if ($PSCmdlet.ShouldProcess(
         $file.Fullname))
      {
         $temp += $file
      }
   }

   # If any files were confirmed, archive them
   if ($temp) {
      $files = $temp

      # Get the fullname to the files
      $FilesFullname = $files | foreach {$_.Fullname}
      
      # Archive the files
      # (7-zip specific code, change if you're using WinZip, GZip, tar, etc.)

      # 7-zip has an idiosyncracy for adding the directory structure into the archive.
      # If we want the path structure of the source files in the output archive,
      # we need to cd to the parent directory and use relative paths for the files.
      
      # -WithPath will only work if all files in the collection are from the same drive!!!
      if ($WithPath) {
         # Create relative paths from the root directory
         # The only colon possible in a directory path is after the drive letter
         $FilesFullname = $FilesFullname | foreach {(($_ -split ":")[1].Substring(1))}
         
         # Change to the drive root
         # Get the drive letter of the first path, and assume the rest are under the same drive.
         Push-Location (($files[0]).PsDrive.root)
      }

      # Archive the files and log to a temporary file
      $tempFileList  = "R:\Temp\filelist.lst"
      $tempFileList  = [IO.Path]::GetTempFileName()
      $tempLogFile   = [IO.Path]::GetTempFileName()
      $wfg           = (Get-Host).PrivateData.WarningForegroundColor
      $wbg           = (Get-Host).PrivateData.WarningBackgroundColor
      $efg           = (Get-Host).PrivateData.ErrorForegroundColor
      $ebg           = (Get-Host).PrivateData.ErrorBackgroundColor

      # Create a file list for 7-zip
      $FilesFullname | Out-File "$tempFileList" -Encoding Ascii
      
      # Now archive the files
      $LASTEXITCODE=0
      try {
         Get-Date | Out-File "$tempLogFile"
         & "$ArchiveProgram" $ArchiveProgramParms "$Archive" `@"$tempFileList" -wR:\Temp |  Out-File "$tempLogFile" -Append
      }
      catch {
         throw "$_"
      }
      $rc=$LASTEXITCODE
      
      Remove-Item $tempFileList -ErrorAction SilentlyContinue

      if ($WithPath) {Pop-Location}

      if ($Log) {
         if ($Append) {Get-Content "$tempLogFile" | Out-File "$Log" -Append}
         else         {Get-Content "$tempLogFile" | Out-File "$Log"}
      }
      else {
         Get-Content "$tempLogFile" | Write-Host
      }

      # what was the exit code?
      switch($rc) {
         0 {
            Get-Content "$tempLogFile" | Write-Host
            Remove-Item "$tempLogFile" -ErrorAction SilentlyContinue
            return $files
            break
         }
         1 {
            Get-Content "$tempLogFile" | Write-Host -ForegroundColor $wfg -BackgroundColor $wbg
            Remove-Item "$tempLogFile" -ErrorAction SilentlyContinue
            Write-Warning "Problems executing 7-Zip, please review log."
            return $null
            break
         }
         default {
            Get-Content "$tempLogFile" | Write-Host -ForegroundColor $efg -BackgroundColor $ebg
            Remove-Item "$tempLogFile" -ErrorAction SilentlyContinue
            Write-Error "Problems executing 7-Zip, please review log."
            return $null
            break
         }
      }
   }
}
#endregion
