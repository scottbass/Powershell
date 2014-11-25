<# HEADER
/*=====================================================================
Program Name            : sas.ps1
Purpose                 : Invoke SAS in batch mode with desired parameters
Powershell Version:     : v2.0
Input Data              : (Optional) List of files in the pipeline
Output Data             : N/A

Originally Written by   : Scott Bass
Date                    : 17SEP2012
Program Version #       : 1.0

=======================================================================

Modification History    :

Programmer              : Scott Bass
Date                    : 26APR2013
Change/reason           : Rewrite, added CmdletBinding and
                          SupportsShouldProcess
Program Version #       : 1.1

=====================================================================*/
#>

<#
.SYNOPSIS
SAS Invocation Script

.DESCRIPTION
Invoke SAS in batch mode with desired parameters.

.PARAMETER  Sysin
SAS program file(s) to execute.

.PARAMETER  Config
SAS configuration file.

.PARAMETER  Lev
SAS level (Lev1, Lev2, Lev9)

.PARAMETER  Logdir
SAS logs directory.

.PARAMETER  Printdir
SAS output directory.  If not specified it defaults to the Logdir directory.

.PARAMETER  Log
SAS log filename.  If not specified the default log filename is <SAS Program Name>_<yyyyMMdd_HHmmss>.log.
Use .Net format strings to change the default log filename:
{0} is the basename of the SAS program file,
{1} is the current datetime, and
{2} is the SAS level.

You can use a format modifier for the datetime to format the datetime portion, for example:
yyyyMMdd_HHmmss, yyMMdd, yyyyMMM, etc.
Use single quotes if you wish to embed runtime Powershell variables in the log filename.
For example, -log '{0}_{1:yyyyMMddHHmmss_${somevariable}.log'

.PARAMETER  Print
SAS output filename.  If not specified it defaults to the same name as the SAS log,
except ending in .lst.
For example:
MySASPgm_20130501.log ==> MySASPgm_20130501.lst, or
MySASPgm.log ==> MySASPgm.lst

.PARAMETER  Options
Additional SAS invocation options.

.PARAMETER  Abort
Abort script when any program ends in either warning or error.
Useful with multiple programs are submitted in one invocation.

.PARAMETER  Quiet
By default SAS parameter values are printed in the console window.
Specify -quiet to switch off this output.

.PARAMETER  Whatif
Echo the SAS invocation command string to the console window,
without actually invoking SAS.

.PARAMETER  Confirm
Asks for confirmation before actually invoking SAS.

.PARAMETER  Verbose
Prints the actually SAS invocation command to the console window.

.EXAMPLE
.\sas.ps1 -sysin SAS_Program_1.sas, or
.\sas SAS_Program_1.sas

Description
-----------
Execute SAS_Program_1.sas in the current directory,
with the default config file, log file, and other options.

.EXAMPLE
.\sas -sysin SAS_Program_3.sas, SAS_Program_2.sas, SAS_Program_1.sas

Description
-----------
Execute SAS_Program_3.sas, SAS_Program_2.sas, and SAS_Program_1.sas in the current directory,
with the default config file, log file, and other options.

.EXAMPLE
.\sas.ps1 (dir etlLoad*.sas), or
gci etlLoad*.sas | .\sas.ps1

Description
-----------
Execute all SAS programs named etlLoad*.sas from the current directory,
with the default config file, log file, and other options.
In the first example, the parentheses are required syntax.
(gci and dir are both aliases for the Get-ChildItem cmdlet).

.EXAMPLE
.\sas -sysin (Get-Content list_of_sas_files.txt), or
gc list_of_sas_files.txt | .\sas.ps1

Description
-----------
Execute all SAS programs defined in the file list_of_sas_files.txt,
in the order defined in the file, with the default config file,
log file, and other options.  In the first example, the parentheses
are required syntax.  (gc is an alias for the Get-Content cmdlet).

.EXAMPLE
.\sas.ps1 -sysin SAS_Program_1.sas -config SAS_Config_File.cfg

Description
-----------
Execute SAS_Program_1.sas in the current directory,
with the specified config file.

.EXAMPLE
.\sas -sysin SAS_Program_1.sas -logdir SAS_Logs_Directory

Description
-----------
Execute SAS_Program_1.sas in the current directory,
with the specified SAS Logs Directory.
By default SAS Listing output will also go to this directory.

.EXAMPLE
.\sas.ps1 -sysin SAS_Program_1.sas -logdir SAS_Logs_Directory -log '{2}_{0}_{1:yyMMdd}.log'

Description
-----------
Execute SAS_Program_1.sas in the current directory,
with the specified SAS Logs Directory, and override the default Log filename.
The format string parameters are:
{0} = Basename of the SAS program file.
{1} = Current datetime.  Format modifiers can be used as shown to format the datetime string.
{2} = SAS level.
So, this log filename might be 'Lev2_MySASProgram_20130501.log'
By default the output filename "follows" the log filename, and so would be
'Lev2_MySASProgram_20130501.lst', and would be written to the Logdir directory.

.EXAMPLE
.\sas.ps1 -sysin SAS_Program_1.sas -options -obs 5 -nosplash -noterminal -batch, or
.\sas SAS_Program_1.sas -obs 5 -nosplash -noterminal -batch -sysparm '"foo bar blah"'

Description
-----------
Execute SAS_Program_1.sas in the current directory,
with the specified command line and SAS options.
All parameters listed on the command line that are not specific script parameters are considered SAS options.
If the SAS options requires quoting (for example -sysparm), the option must be "double quoted".
For example:
-sysparm '''foo bar blah''', or
-sysparm """foo bar blah"""      (three single (or double) quotes)
-sysparm "'foo bar blah'", or
-sysparm '"foo bar blah"'        (quotes within quotes)
-sysparm `'foo bar blah`', or
-sysparm `"foo bar blah`"        (escaped single (or double) quotes (the backtick is the escape character)

.EXAMPLE
.\sas -sysin Success.sas, Warning.sas, Error.sas, Success2.sas -abort warning

Description
-----------
Abort the script when a job ends with a warning or greater.
In this example, the script would halt after Warning.sas

.EXAMPLE
.\sas -sysin Success.sas, Warning.sas, Error.sas, Success2.sas -abort error

Description
-----------
Abort the script when a job ends with an error or greater.
In this example, the script would halt after Error.sas

.EXAMPLE
.\sas -sysin SAS_Program_1.sas -quiet

Description
-----------
Execute SAS_Program_1.sas in the current directory,
suppressing the echo of SAS parameters to the console window.

.EXAMPLE
dir *.sas | .\sas -whatif

Description
-----------
Execute every SAS program in the current directory,
echoing the SAS command invocation string to the console window,
without actually invoking SAS.

.EXAMPLE
dir *.sas | .\sas -confirm

Description
-----------
Execute every SAS program in the current directory,
asking for confirmation before actually invoking SAS.

.NOTES
This script uses a number of internal variables for certain functionality.  The variable names all begin with "sas".
Here is the list of variables, their usage, and the script default values:

VARIABLE       USAGE                         DEFAULT
========       =====                         =======
Global:
-------
$saslev        SAS Level (Lev1, Lev2, Lev9)  Lev1
$sasexe        SAS Invocation Command        D:\Program Files\SASHome\SASFoundation\9.3\sas.exe
$sasconfig     SAS Configuration File        E:\SAS\Config\$saslev\SASApp\sasv9.cfg
$saslogdir     SAS Logs Directory            E:\Logs\$saslev\BatchServer\SASApp
$sasprintdir   SAS Output Directory          E:\Print\$saslev\BatchServer\SASApp
$sasoptions    SAS Options                   -metaautoresources "SASApp"

Per-Program:
------------
$saslog        SAS Log Filename           "{0}_{1:yyyyMMdd_HHmmss}.log"
$sasprint      SAS Print Filename         ($log -replace '.log$','')+".lst"

By default, the SAS print file "follows" the name of the SAS log file.

In other words, if the SAS log file is MySASProgram_20130501_123456.log,
then the SAS print file is MySASProgram_20130501_123456.log.

And if the SAS log filename is MySASProgram_20130501.log,
then the SAS print file is MySASProgram_20130501.log.

Both of these defaults can be overridden with the -log and -print command line options.

The Global variables have a single value for the life of the script invocation.
The Per-Program variables usually have a different value per each SAS program,
since by default they use replaceable format strings ({0}, {1}, and {2}).

The replaceable format strings are:

{0}:  Basename (name without extension) of the SAS program
{1}:  Current datetime
{2}:  SAS level

Format modifiers can be used with the current datetime to control the elements returned.

You can set the value of one of these SAS variables in the console window to override the default value.
This can be very useful during program development.

For example, if you set this in your console window:

$saslev="Lev2"
$saslog="{0}.log"

Then this script will run in SAS Level 2, and will not create a time stamped log,
until these variables are changed or deleted, unless overridden by a command line option.

Use the command del variable:saslog to delete the $saslog variable,
or del variable:sas* to delete all variables beginning with "sas".

.LINK
http://ss64.com/ps/syntax-dateformats.html
http://msdn.microsoft.com/en-us/library/8kb3ddd4.aspx

#>

#region Parameters
[CmdletBinding(SupportsShouldProcess=$true)]
param(
   [Alias("Fullname")]
   [Parameter(
      Position=0,
      ValueFromPipeline=$true,
      ValueFromPipelineByPropertyName=$true
   )]
   [Object[]]$sysin
   ,
   [ValidateSet("Lev1","Lev2","Lev9")]
   [String]$lev
   ,
   [String]$exe
   ,
   [String]$config
   ,
   [String]$logdir
   ,
   [String]$printdir=$logdir
   ,
   [String]$log="{0}_{1:yyyyMMdd_HHmmss}.log"
   ,
   [String]$print=($log -replace '.log$','')+".lst"
   ,
   [ValidateSet("WARNING","ERROR")]
   [String]$abort
   ,
   [Switch]$quiet
   ,
   [Switch]$noautoresources
   ,
   [Parameter(ValueFromRemainingArguments=$true)]
   [String[]]$options
)
#endregion

#region Main
BEGIN {
   #=== This section is called once, at the beginning of the script ===#

   Set-StrictMode -Version Latest
   $ErrorActionPreference="Stop"
   $defaults=@{}
   $rtncodes=@()
   $sasoptions=@()
   $sasrc=$null

   # set hardcoded default values here
   $defaults.Add("saslev",       'Lev1')
   $defaults.Add("sasexe",       'D:\Program Files\SASHome\SASFoundation\9.3\sas.exe')
   $defaults.Add("sasconfig",    'E:\SAS\Config\$saslev\SASApp\sasv9.cfg')
   $defaults.Add("saslogdir",    'E:\Logs\$saslev\BatchServer\SASApp')
   $defaults.Add("sasprintdir",  'E:\Print\$saslev\BatchServer\SASApp')
   $defaults.Add("sasoptions",   '-metaautoresources "SASApp"')

   # sas* variables set in the calling process will be used in this script
   # otherwise create and initialize the variables to null
   if (!(Test-Path variable:saslev))               {$saslev       = $null}
   if (!(Test-Path variable:sasexe))               {$sasexe       = $null}
   if (!(Test-Path variable:sasconfig))            {$sasconfig    = $null}
   if (!(Test-Path variable:saslogdir))            {$saslogdir    = $null}
   if (!(Test-Path variable:sasprintdir))          {$sasprintdir  = $null}
   if (!(Test-Path variable:sasoptions))           {$sasoptions   = $null}
                                                    $sasenv       = $null

   # parameters set on the command line will override any external variables
   if ($PSBoundParameters.ContainsKey("exe"))      {$sasexe       = $exe}
   if ($PSBoundParameters.ContainsKey("config"))   {$sasconfig    = $config}
   if ($PSBoundParameters.ContainsKey("lev"))      {$saslev       = $lev}
   if ($PSBoundParameters.ContainsKey("logdir"))   {$saslogdir    = $logdir}
   if ($PSBoundParameters.ContainsKey("printdir")) {$sasprintdir  = $printdir}
   if ($options)                                   {$sasoptions   = $options}

   # if the variable is still not set by this point, set defaults for certain variables
   if (!($saslev))                                 {$saslev       = $defaults.saslev}

   # make $saslev Proper Case (first letter capitalized)
   $saslev = $saslev.Replace("l","L");

   switch ($saslev) {
      "Lev1" {$sasenv = "prod";  break}
      "Lev2" {$sasenv = "test";  break}
      "Lev9" {$sasenv = "admin"; break}
   }

   if (!($sasexe))                                 {$sasexe       = $defaults.sasexe}
   if (!($sasconfig))                              {$sasconfig    = $defaults.sasconfig}
   if (!($saslogdir))                              {$saslogdir    = $defaults.saslogdir}
   if (!($sasprintdir))                            {$sasprintdir  = $defaults.sasprintdir}

   # add default options to the end of any specified options
   if (!($noautoresources))                        {$sasoptions += $defaults.sasoptions}

   # resolve any embedded variables from external variable definition
   # (usually $saslev but could be anything)
   if ($saslev)      {Try {$saslev      = Invoke-Expression "Write-Output `"$saslev`""}      Catch {}}
   if ($sasenv)      {Try {$sasenv      = Invoke-Expression "Write-Output `"$sasenv`""}      Catch {}}
   if ($sasexe)      {Try {$sasexe      = Invoke-Expression "Write-Output `"$sasexe`""}      Catch {}}
   if ($sasconfig)   {Try {$sasconfig   = Invoke-Expression "Write-Output `"$sasconfig`""}   Catch {}}
   if ($saslogdir)   {Try {$saslogdir   = Invoke-Expression "Write-Output `"$saslogdir`""}   Catch {}}
   if ($sasprintdir) {Try {$sasprintdir = Invoke-Expression "Write-Output `"$sasprintdir`""} Catch {}}

   # validate values
   # saslev must be Lev1, Lev2, or Lev9
   if ("Lev1","Lev2","Lev9" -notcontains $saslev) {
      Throw "ERROR: Level parameter $saslev must be Lev1, Lev2, or Lev9 (case-insensitive)."
   }

   # sasconfig must exist and must be a file
   if (!(Test-Path "$sasconfig" -PathType Leaf)) {
      Throw "ERROR: SAS Configuration File $sasconfig does not exist."
   }

   # saslogdir must exist and must be a directory
   if (!(Test-Path "$saslogdir" -PathType Container)) {
      Throw "ERROR: SAS Log Directory $saslogdir either does not exist or is not a directory."
   }

   # sasprintdir must exist and must be a directory
   if (!(Test-Path "$sasprintdir" -PathType Container)) {
      Throw "ERROR: SAS Print Directory $sasprintdir either does not exist or is not a directory."
   }

   # define utility functions
   Function Print-Vars
   {
      # list the SAS variables
      $name= @{Label="Name";  Expression={$_.Name};  Width=30}
      $value=@{Label="Value"; Expression={$_.Value}; Width=120}

      # convert the sasoptions array to a string for better output
      $local:sasoptions=$sasoptions -join " "

      Push-Location variable:
      Get-Item saslev,sasenv,sasexe,sasconfig,saspgm,saslogfull,sasprintfull,sasoptions | Format-Table $name,$value  # print in this order
      Write-Output "`n"
      Pop-Location
   }

   Function Invoke-SAS {
      # Start SAS
      $p=Start-Process "$sasexe" `
            -ArgumentList "-sysin ""$saspgm"" -config ""$sasconfig"" -log ""$saslogfull"" -print ""$sasprintfull"" $sasoptions " `
            -Wait `
            -NoNewWindow `
            -PassThru

      $script:sasrc=$p.ExitCode
   }

   Function Print-SASResultMsg {
      Param(
         [int] $rc
      )
      if ($rc -eq $null) {return}

      switch ($rc) {
         0 {Write-Output "SAS Ended Successfully"}
         1 {Write-Output "SAS Ended With Warnings"}
         2 {Write-Output "SAS Ended With Errors"}
         3 {Write-Output "User issued the ABORT statement"}
         4 {Write-Output "User issued the ABORT RETURN statement"}
         5 {Write-Output "User issued the ABORT ABEND statement"}
         6 {Write-Output "SAS internal error"}
         default {Write-Output "User specified RETURN code: $rc"}
      }
      Write-Output "`n"
   }

   Function Set-SASReturnCode {
      Param(
         [int] $rc
      )
      if ($rc -eq $null -or $rc -eq 0) {return}

      <#
      Certain fatal operating system errors return errorcode=1.
      A SAS warning also returns errorcode=1.

      We want JAMS processing to halt on the fatal o/s error,
      but continue on the SAS warning.

      When we run multiple jobs in one invocation of this script,
      we also want the severity of the SAS return code to be maintained
      so that the maximum return code is returned to the operating system.

      The workaround/solution is to offset the SAS return code by a constant,
      then configure JAMS to treat the SAS warning as a warning, not error.

      In this case, we will add 1000 to the SAS return code, then configure
      JAMS to treat 1001 as a warning condition.
      #>

      $script:sasrc=1000+$rc
   }

   Function Print-SASErrorsOrWarnings {
      Param(
         [int] $rc
      )
      if ($rc -eq $null) {return}

      # post-process log file if errors or warnings occurred
      if ($rc -ne 0) {
         if (Test-Path "$saslogfull") {
            Select-String `
               -path "$saslogfull" `
               -pattern "ERROR:|WARNING:" `
               -context 1,1 `
               | ForEach-Object {$_.LineNumber.ToString() + ": " + $_.Line}
            Write-Output "`n"
         }
      }
   }
}

PROCESS {
   #=== This section is called for each object in the pipeline ===#

   foreach($files in $sysin) {
      # resolve any embedded variables
      # (usually $saslev but could be anything)
      Try {$files = Invoke-Expression "Write-Output `"$files`""}  Catch {}

      # convert to a file system object
      # (use Get-ChildItem, $files could contain a filename wildcard resulting in many files)
      $programs = Get-ChildItem "$files" -ErrorAction SilentlyContinue

      # if no files found then abort
      if (!($programs)) {
         Throw "ERROR: No SAS Program Files found matching $files."
      }

      foreach($saspgm in $programs) {
         # saspgm is required and must exist
         if (! $saspgm) {
            Throw "ERROR: SAS Program File is required."
         }
         if (!(Test-Path $saspgm -PathType Leaf)) {
            Throw "ERROR: SAS Program File $saspgm.FullName does not exist."
         }

         # snapshot the current datetime
         $datetime=Get-Date

         # derive log and print filenames
         # if a command line parameter was specified, use it
         # else if a global variable was specfied, use it
         # otherwise use the defaults
         # derive the log first, then derive the print file

         # log file
         if ($PSBoundParameters.ContainsKey("Log")) {
            # nothing, log value is set via the command line parameter
         }
         elseif (Get-Variable -Name saslog -Scope Global -ErrorAction SilentlyContinue) {
            $log = $global:saslog
         }
         else {
            # nothing, use the default value set via the parameters
         }

         # resolve formatting directives ( {0}, {1}, {2} )
         $local:saslog=
            Try {
               $log -f $saspgm.BaseName, $datetime, $saslev
            } Catch {
               Throw "ERROR: Unable to set the log file name using format string $log."
            }

         # print file
         if ($PSBoundParameters.ContainsKey("Print")) {
            # nothing, print value is set via the command line parameter
         }
         elseif (Get-Variable -Name sasprint -Scope Global -ErrorAction SilentlyContinue) {
            $print = $global:sasprint
         }
         else {
            # rederive based on the default functionality
            # (need to keep this code in sync with the parameter default above)
            $print=($log -replace '.log$','')+".lst"
         }

         # resolve formatting directives ( {0}, {1}, {2} )
         $local:sasprint=
            Try {
               $print -f $saspgm.BaseName, $datetime, $saslev
            } Catch {
               Throw "ERROR: Unable to set the print file name using format string $print."
            }

         # full path to SAS log and print files
         $saslogfull    = $saslogdir,  $saslog   -join "\"
         $sasprintfull  = $sasprintdir,$sasprint -join "\"

         # resolve any embedded variables
         Try {$saslogfull  = Invoke-Expression "Write-Output `"$saslogfull`""}     Catch {}
         Try {$sasprintfull= Invoke-Expression "Write-Output `"$sasprintfull`""}   Catch {}

         # "`n$sasexe `n-sysin ""$saspgm"" `n-config ""$sasconfig"" `n-log ""$saslogfull"" `n-print ""$sasprintfull"" `n$sasoptions","Invoke SAS"
         # "$sasexe -sysin ""$saspgm"" -config ""$sasconfig"" -log ""$saslogfull"" -print ""$sasprintfull"" $sasoptions","Invoke SAS"

         # Invoke SAS unless -whatif was specified (-verbose and -confirm are additional options)
         if ($PSCmdlet.ShouldProcess(
            "`n$sasexe `n-sysin ""$saspgm"" `n-config ""$sasconfig"" `n-log ""$saslogfull"" `n-print ""$sasprintfull"" `n$sasoptions","Invoke SAS"
         ))
         {
            $local:ConfirmPreference="High"
            Try {
               if (! $quiet) {Print-Vars}
               Invoke-SAS
               if (! $quiet) {
                  Print-SASResultMsg $script:sasrc
                  Print-SASErrorsOrWarnings $script:sasrc
               }
               Set-SASReturnCode $script:sasrc

               # add the return code to the array of return codes
               $rtncodes+=$script:sasrc

               # if abort was specified, exit the script on warning or error
               switch -wildcard ($abort) {
                  W* {
                     if ($script:sasrc -ge 1001) {
                        $global:lastexitcode=($rtncodes | measure-object -max).maximum
                        Throw "ABORT: Abort on "+$abort.toupper()+" was specified.  Script halted."
                     }
                  }
                  E* {
                     if ($script:sasrc -ge 1002) {
                        $global:lastexitcode=($rtncodes | measure-object -max).maximum
                        Throw "ABORT: Abort on "+$abort.toupper()+" was specified.  Script halted."
                     }
                  }
               }
            }
            Catch {
               Write-Error "$_"
            }
         }
      }
   }
}

END {
   #=== This section is called once, at the end of the script ===#

   # get the maximum return code
   $global:lastexitcode=($rtncodes | measure-object -max).maximum
}
#endregion
