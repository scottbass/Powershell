<# HEADER
/*=====================================================================
Program Name            : Remove-Files.ps1
Purpose                 : Function to remove files by various criteria.
Powershell Version:     : v2.0
Input Data              : (Optional) List of files in the pipeline
Output Data             : N/A

Originally Written by   : Scott Bass
Date                    : 01MAR2013
Program Version #       : 1.0

=======================================================================

Modification History    :

=====================================================================*/
#>

<#
.SYNOPSIS
Script to remove files by various criteria.

.DESCRIPTION
Remove-Files is used to remove files matching various selection criteria.

The file selection criteria is implemented by the Select-Files function.
This function must be loaded in order for this script to work.

The selection criteria is documented in the Select-Files function.
Type "help Select-Files" for more information.

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

.PARAMETER  Type
Lists only files of the specified type. Valid values are "file" and "dir".

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

See examples below.

.EXAMPLE
Remove-Files .
Archive-Files -path .

Description
-----------
Remove files in the current directory.

.EXAMPLE
Remove-Files -path C:\Temp -filter *.txt

Description
-----------
Remove *.txt files in the C:\Temp directory.

.EXAMPLE
Remove-Files -path C:\Temp\Temp1,C:\Temp\Temp2 -exclude *.7z,*.log -recurse

Description
-----------
Remove files in the C:\Temp\Temp1 and C:\Temp\Temp2 directories, including all
sub-directories, excluding *.7z and *.log files.

.EXAMPLE
Remove-Files -path C:\Temp -whatif

Description
-----------
Display what *would* happen (-whatif) in each of these examples.
Does not actually remove the files.

.EXAMPLE
Remove-Files -path C:\Temp -confirm

Description
-----------
Lists the files that will be removed, giving you the opportunity to confirm the
operation (-confirm).

.EXAMPLE
Select-Files -path C:\Temp -size 10mb | Remove-Files -confirm

Description
-----------
Use Select-Files to select files in C:\Temp larger than 10mb, and pipe those files
to Remove-Files, giving you the opportunity to confirm the
operation (-confirm).

.EXAMPLE
Select-Files -path C:\Temp -filter *.txt -recurse | Archive-Files -archive C:\Temp\Archive.7z -withpath | Remove-Files

Description
-----------
Use Select-Files to select *.txt files in C:\Temp and all sub-directories (-recurse),
piping those files select to the Archive-Files script for archiving in the C:\Temp\Archive.7z file.
If the archive is successful (no errors thrown, 0 return code, processed files returned to the pipeline),
then remove those files returned in the pipeline.
#>

#region Parameters
[CmdletBinding(SupportsShouldProcess=$true)]
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
   [Parameter(Position = 1)]
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
   [ValidateSet("file","dir")]
   [String]$Type
   ,
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
   [Parameter(ValueFromRemainingArguments=$true)]
   $Dummy
)
#endregion

#region Main
begin {
   $ErrorActionPreference = "Stop"

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
   if (-not $files) {Write-Warning "There are no files to process."; exit $null}

   # Now filter the files
   # Pass the script parameters to Select-Files by splatting the $PSBoundParameters collection.
   $PSBoundParameters.Item("Path")=$files
   $files = Select-Files @PSBoundParameters

   # If there are no files to process then abort.
   if (-not $files) {Write-Warning "There are no files to process."; exit $null}

   # If -whatif was specified print what we would do
   # If -confirm was specified allow the user to select which files to process
   foreach ($file in $files) {
      if ($PSCmdlet.ShouldProcess(
         $file.Fullname))
      {
         $local:ConfirmPreference="High"
         try {
            Remove-Item $file -ErrorAction "Continue"
         }
         catch {
            Write-Error "$_"
         }
      }
   }
}
#endregion
