function Select-Files {
<# HEADER
/*=====================================================================
Program Name            : Select-Files.ps1
Purpose                 : Function to list files by various criteria.
Powershell Version:     : v2.0
Input Data              : N/A
Output Data             : Array of file objects/Array of file names

Originally Written by   : Scott Bass
Date                    : 01MAR2013
Program Version #       : 1.0

=======================================================================

Modification History    :

Programmer              : Scott Bass
Date                    : 03APR2013
Change/reason           : Added owner parameter
Program Version #       : 1.1

=====================================================================*/
#>

<#
.SYNOPSIS
Script to list files by various criteria.

.DESCRIPTION
Select-Files is used to list files matching various selection criteria.

Much of the script is just a "wrapper" around the functionality of Get-ChildItem
(enter "help Get-ChildItem" for more details), although additional filtering
criteria is also available.

The selection criteria is cumulative, so only those files that match all
criteria are returned.

The results can be assigned to a variable or piped to other cmdlets or scripts.

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

.Example
Select-Files

Description
-----------
List files and sub-directories in the current directory (same results as gci, dir, ls, etc.)

.Example
Select-Files C:\Temp
Select-Files -path C:\Temp

Description
-----------
List files and sub-directories in the C:\Temp directory

.Example
Select-Files C:\Temp .txt,.xml,.zip
Select-Files -path C:\Temp .txt,.xml,.zip
Select-Files -path C:\Temp -ext .txt,.xml,.zip
Select-Files -path C:\Temp -extension .txt,.xml,.zip

Description
-----------
List .txt, .xml, and .zip files in the C:\Temp directory

.Example
Select-Files -path C:\Temp,C:\Temp2,C:\Temp3 -filter *.txt

Description
-----------
List .txt files in the C:\Temp, C:\Temp2, and C:\Temp3 directories

Only one filter can be specified.

.Example
Select-Files -path C:\Temp\*,C:\Temp2\* -include *.txt,*.xls*,*.zip

Description
-----------
List .txt, Excel workbooks (.xls, .xlsx, etc.), and .zip files
in the C:\Temp and C:\Temp2 directories.

Note the trailing asterisk on the path(s).  If the asterisk is not specified,
no files will be returned.

.Example
Select-Files -path C:\Temp -exclude *.txt,*.xml

Description
-----------
List all files in the C:\Temp directory except .txt and .xml files.

Note no trailing asterisk on the path.  If the asterisk is specified,
files in sub-directories of the -path will be returned (except for
the files excluded by the -exclude list).

.Example
Select-Files -path C:\Temp -filter a* -type file

Description
-----------
List only files (no sub-directories) beginning with "a" in the C:\Temp directory.

.Example
Select-Files -path C:\Temp -older 01sep12 | sort -property LastWriteTime
Select-Files -path C:\Temp -older "09/01/12 00:00:00" | sort -property LastWriteTime

Description
-----------
List files in C:\Temp older than 01-Sep-2012 00:00:00

.Example
Select-Files -path C:\Temp -newer 01-Jan-2013 | sort -property LastWriteTime
Select-Files -path C:\Temp -newer "01-Jan-2013 00:00:00" | sort -property LastWriteTime

Description
-----------
List files in C:\Temp newer than 01-Jan-2013 00:00:00

.Example
Select-Files -path C:\Temp -between 01jan13,01FEB13 | sort -property LastWriteTime
Select-Files -path C:\Temp -between 01jan13,"31-Jan-2013 23:59:59" | sort -property LastWriteTime

Description
-----------
List files in C:\Temp between either 01-Jan-13 00:00:00 and 01-Feb-13 00:00:00, or
List files in C:\Temp between either 01-Jan-13 00:00:00 and 31-Jan-13 23:59:59 (inclusive)

.Example
Select-Files -path C:\Temp -type file -size 1mb | sort -property Length
Select-Files -path C:\Temp -type file -size (-5kb) | sort -property Length

Description
-----------
List files (no sub-directories) in C:\Temp, then within that list return the files larger than 1MB, or
List files (no sub-directories) in C:\Temp, then within that list return the files smaller than 5KB

Note that negative size specifications using the Powershell mnemonics must be enclosed in parentheses.

.Example
Select-Files -path C:\Temp -filter *.txt -age 30 | sort -property LastWriteTime
Select-Files -path C:\Temp -filter *.txt -age -30 | sort -property LastWriteTime

Description
-----------
List .txt files in C:\Temp older than 30 days ago (from today's date), or
List .txt files in C:\Temp newer than 30 days ago (from today's date)

.Example
Select-Files -path C:\Temp -num 5  | sort -property LastWriteTime
Select-Files -path C:\Temp -num -5 | sort -property LastWriteTime

Description
-----------
List files in C:\Temp, then within that list return the 5 oldest files, or
List files in C:\Temp, then within that list return the 5 newest files

.Example
Select-Files -path C:\Temp -size 1 -num 10  | sort -property Length
Select-Files -path C:\Temp -size 1 -num -10 | sort -property Length

Description
-----------
List files in C:\Temp larger than 1 byte, then within that list return the 10 largest files, or
List files in C:\Temp larger than 1 byte, then within that list return the 10 smallest files

.Example
Select-Files -path C:\Temp -size 10MB -num 10 | sort -property Length

Description
-----------
List the Top 10 files in C:\Temp larger than 10MB

.Example
Select-Files -path C:\Temp -filter *.txt -older 01jan13 -whichdate created | sort -property CreationTime | Format-Table CreationTime, Name -auto

Description
-----------
List .txt files in C:\Temp whose creation date is older than 01-Jan-2013

.Example
Select-Files -path C:\Temp -filter *.txt -newer 01jan13 -whichdate accessed | sort -property LastAccessTime | Format-Table LastAccessTime, Name -auto

Description
-----------
List .txt files in C:\Temp whose last access date is newer than 01-Jan-2013

.Example
Select-Files -path C:\Temp -type file -recurse -name -size 1MB

Description
-----------
List all files (-type file) in C:\Temp, and all sub-directories (-recurse),
returning only the filenames (array of strings) rather than the file system objects
(array of System.IO.FileInfo objects)

.Example
$exp={(Get-Date $_.$DateType -f MM) -eq "09"}, OR
$exp={(Get-Date $_.$DateType -f MMM) -ieq "SeP"},

Select-Files -path C:\Temp -filter *.txt -whichdate modified -ScriptBlock $exp | sort -property LastAccessTime

Description
-----------
List .txt files in C:\Temp, then within that list return those files whose
LastWriteTime ("modified") was in September (of any year).

.Example
$exp={(Get-Date $_.$DateType -f MMM) -eq ((Get-Date).AddMonths(-1) | Get-Date -f MMM)}

Select-Files -path C:\Temp -filter *.txt -whichdate created -ScriptBlock $exp | sort -property CreationTime | Format-Table CreationTime, Name -auto

Description
-----------
List .txt files in C:\Temp, then within that list return those files whose
CreationTime ("created") was last month.

.Example
$exp={$QTR[(Get-Date $_.$DateType -f MMM)] -eq "Q4"}

Select-Files -path C:\Temp -filter *.log -recurse -type file -between 01jan12,01jan13 -script $exp | sort -property LastWriteTime

Description
-----------
List .log files in C:\Temp, and all sub-directories (-recurse), returning only files (-type file),
whose LastWriteTime (the script default for -whichdate) is between 01-Jan-2012 and 01-Jan-2013 (i.e. during 2012).
Then, within that list, return those files whose LastWriteTime was during Q4.

Note that hash key lookup is case-insensitive; $QTR["jan"], $QTR["Jan"], and $QTR["JAN"]
all return the same results.

.Example
Select-Files -path C:\Temp -filter *.txt -verbose -whatif

Description
-----------
Echo the script parameters to the console (-verbose) and suppress the actual listing of the files (-whatif).
Useful during debugging, although rarely used since this script is not destructive.

.Link
http://ss64.com/ps/syntax-dateformats.html
 http://ss64.com/bash/date.html#format
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
   [Int64]$Size
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
   
   ### Additional attributes parameters
   [Switch]$owner
   ,
   
   ### Catchall parameter
   [Parameter(ValueFromRemainingArguments=$true)]
   $Dummy
)
#endregion

#region Functions
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
   # Validation of -between parameter
   if ($Between) {
      if ($Between.Length -ne 2) {throw "The -between parameter requires both a start and end date."}
   }

   # Create a $QTR hash table for filtering files by quarter (via advanced expressions)
   $local:QTR=@{}
   $QTR.Add("Jan","Q1")
   $QTR.Add("Feb","Q1")
   $QTR.Add("Mar","Q1")
   $QTR.Add("Apr","Q2")
   $QTR.Add("May","Q2")
   $QTR.Add("Jun","Q2")
   $QTR.Add("Jul","Q3")
   $QTR.Add("Aug","Q3")
   $QTR.Add("Sep","Q3")
   $QTR.Add("Oct","Q4")
   $QTR.Add("Nov","Q4")
   $QTR.Add("Dec","Q4")

   # First process all the "built-in" Get-ChildItem parameters for better performance,
   # except -Name, we will process that at the end

   # Build Get-ChildItem command string
   $cmd="Get-ChildItem"
   if ($files) {
      $files_string=($files | foreach {"'{0}'" -f $_}) -join ","
      $cmd += " -path $files_string"
   }
   if ($Filter)     {$cmd += " -filter $filter"}
   if ($Include)    {$cmd += " -include " + ($include -join ",")}
   if ($Exclude)    {$cmd += " -exclude " + ($exclude -join ",")}
   if ($Recurse)    {$cmd += " -recurse"}

   # Capture list of files
   try {
      $files=Invoke-Expression $cmd
   }
   catch {
      throw "$_"
   }
   
   # Convert the -whichdate parameter to the actual file attribute
   switch($WhichDate) {
      "created"   {$DateType="CreationTime";break}
      "accessed"  {$DateType="LastAccessTime";break}
      "modified"  {$DateType="LastWriteTime";break}
   }

   # Convert the -type parameter to the actual file attribute
   switch ($Type) {
      "file"      {$FileType=$false;break}
      "dir"       {$FileType=$true;break}
   }

   # Now apply additional filtering.  All filtering is cumulative.
   if ($Extension) {$files = $files | Where {$Extension -icontains $_.extension}}
   if ($Type)      {$files = $files | Where {$_.PsIsContainer -eq $FileType}}
   if ($Older)     {$files = $files | Where {$_.$DateType -lt $Older}}
   if ($Newer)     {$files = $files | Where {$_.$DateType -gt $Newer}}

   # If -between was specified, return files where $DateType is between the start and end dates
   # Note: best practice is not to specify -older or -newer if -between is specified, but no error checking is done
   # You get the list you asked for!
   if ($Between) {
      $start=Get-Date $Between[0]
      $end=  Get-Date $Between[1]
                    $files = $files | Where {($start -le $_.$DateType) -and ($_.$DateType -le $end)}
   }

   # Note: Integer parameters are initialized to zero, even when not specified

   # If -size is positive, return files larger than $size
   # If -size is negative, return files smaller than $size
   if ($Size -ne 0) {
      if ($Size -ge 0)
                   {$files = $files | Where {$_.Length -gt ([math]::Abs($Size))}}
      else
                   {$files = $files | Where {$_.Length -lt ([math]::Abs($Size))}}
   }

   # Age is based on today's date minus $age number of days
   # If -age is positive, return files older than (today-age) (based on $DateType)
   # If -age is negative, return files newer than (today-age) (based on $DateType)
   if ($Age -ne 0) {
      # Derive age as a Date object
      $AgeDate = (Get-Date).AddDays(-([Math]::Abs($Age)))

      if ($Age -ge 0)
                   {$files = $files | Where {$_.$DateType -lt $AgeDate}}
      else
                   {$files = $files | Where {$_.$DateType -gt $AgeDate}}
   }

   # If -num is positive, return $num <oldest|largest> files (based on $DateType)
   # If -num is negative, return $num <newest|smallest> files (based on $DateType)
   if ($Num -ne 0) {
      if ($Num -ge 0) {
         if ($size -ne 0) # At this point, the value of -size is irrelevant, just that it was specified
                   {$files = $files | Sort-Object -Property Length    | select -Last  ([math]::Abs($Num)) | Sort-Object} # Largest
         else      {$files = $files | Sort-Object -Property $DateType | select -First ([math]::Abs($Num)) | Sort-Object} # Oldest
      }
      else {
         if ($size -ne 0)
                   {$files = $files | Sort-Object -Property Length    | select -First ([math]::Abs($Num)) | Sort-Object} # Smallest
         else      {$files = $files | Sort-Object -Property $DateType | select -Last  ([math]::Abs($Num)) | Sort-Object} # Youngest
      }
   }

   # If an advanced expression was specified, further filter the list
   if ($ScriptBlock -ne $null) {
      try          {$files = $files | Where {Invoke-Expression $ScriptBlock.ToString()}}
      catch        {throw "Error invoking ScriptBlock: $ScriptBlock"}
   }

   # If -Name was specified, return full pathnames to the pipeline
   # If -Owner was specified, return custom objects containing the file object and file owner object
   # Otherwise return file objects
   if ($Name) {
      $files | ForEach {$_.Fullname}
   }
   elseif ($Owner) {
      $files | ForEach {
         $object = New-Object -TypeName PSObject
         $object | Add-Member -MemberType NoteProperty -Name File -Value $_
         $object | Add-Member -MemberType NoteProperty -Name Owner -Value ($_.GetAccessControl()).Owner
         $object
      }
   }
   else {
      $files
   }
}
#endregion
}
