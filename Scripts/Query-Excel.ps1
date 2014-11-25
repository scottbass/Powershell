<# HEADER
/*=====================================================================
Program Name            : Query-Excel.ps1
Purpose                 : Execute query against Excel worksheet
Powershell Version:     : v2.0
Input Data              : N/A
Output Data             : N/A

Originally Written by   : Scott Bass
Date                    : 04OCT2013
Program Version #       : 1.0

=======================================================================

Modification History    :

Programmer              : Scott Bass
Date                    : 29JAN2014
Change/reason           : Added delimiter parameter
Program Version #       : 1.1

Programmer              : Scott Bass
Date                    : 13FEB2014
Change/reason           : Added culture (date formatting) code
Program Version #       : 1.2

=====================================================================*/

/*---------------------------------------------------------------------

THIS SCRIPT MUST RUN UNDER x86 (32-bit) POWERSHELL SINCE WE ARE USING
32-BIT MICROSOFT OFFICE.  ONLY THE x86 OLEDB PROVIDER IS INSTALLED!!!

The format of the SQL Query MUST be:

select [top x] [column names | *] from [Sheet1$]

Yes, the brackets and trailing $ sign are REQUIRED!

These queries will fail with an error (read the error message!):

select * from [Sheet1]
"The Microsoft Access database engine could not find the object 'Sheet1'.

select * from Sheet1$
"Syntax error in FROM clause."

=======================================================================

NOTE:

Excel is notoriously problemmatic when dealing with columns containing
mixed data types.

Furthermore, IMO the Microsoft.ACE.OLEDB.12.0 is rather buggy in this
regard.

This problem manifests itself when reading character data types in
rows greater than the TypeGuessRows registry setting.

These issues are documented in many places on the Internet, here are
two hits:

http://blogs.lessthandot.com/index.php/datamgmt/dbprogramming/mssqlserver/what-s-the-deal-with/
http://yoursandmyideas.wordpress.com/2011/02/05/how-to-read-or-write-excel-file-using-ace-oledb-data-provider/

(At least) two workarounds are possible:

1.A) Specify HDR=NO on the connection string.  This causes the columns
to be named as F1, F2, F3, etc, and the header row (which is character)
is the first row of data.
1.B) IMEX=1 must be specified on the connection string.  This allows
mixed data types to be read as character, which is fine since we are
(usually) creating CSV files.
1.C) Use the Select-Object cmdlet to skip the header row so only data
is returned.

The downside of this approach is you cannot specify
select <var1, var2, var3, etc> from [SheetName$]
as the query string, since the columns are named F1, F2, F3, etc.

2) Edit the TypeGuessRows registry setting, changing from the default
of 8 to 0.  This causes Excel and the OLEDB driver to scan 16384 rows
to determine the data type.

The downside of this approach is if the desired data type (usually
character) does not appear in the first 16384 rows, the column will be
cast as numeric and the character data will be set to null.

I have chosen option 2.  The code to implement option 1 is still
included but is commented out.

---------------------------------------------------------------------*/
#>

<#
.SYNOPSIS
Query Excel Worksheet

.DESCRIPTION
Execute a query against an Excel Worksheet

.PARAMETER  SQLQuery
SQL Query to execute

.PARAMETER  Path
Path to Excel Workbook

.PARAMETER  Csv
Output as CSV?  If no, the Dataset Table object is returned to the pipeline

.PARAMETER  Delimiter
Override default comma delimiter.  Only used when output as CSV.

.PARAMETER  Whatif
Echos the SQL query information without actually executing it.

.PARAMETER  Confirm
Asks for confirmation before actually executing the query.

.PARAMETER  Verbose
Prints the SQL query to the console window as it executes it.

.EXAMPLE
.\Query-Excel.ps1 C:\Temp\Temp.xlsx "select top 10 * from [Sheet1$]" -csv

Description
-----------
Queries the specified Excel workbook and worksheet with the specified query, outputting data as CSV.

.EXAMPLE
.\Query-Excel.ps1 C:\Temp\Temp.xlsx "select top 10 * from [Sheet1$]" -csv -delimiter "|"

Description
-----------
Queries the specified Excel workbook and worksheet with the specified query, outputting data as a pipe separated file.

.EXAMPLE
.\Query-Excel.ps1 C:\Temp\Temp.xlsx "select claimnum,covno,suffix from [Sheet1$] where claimnum like '2061301%' order by covno,suffix" -csv

Description
-----------
Queries the specified Excel workbook and worksheet with the specified query, outputting data as CSV

.EXAMPLE
.\Query-Excel.ps1 C:\Temp\Temp.xlsx "select * from Sheet1" -csv:$false

Description
-----------
Queries the specified Excel workbook and worksheet with the specified query,
returning the Object Table to the pipeline

#>

#region Parameters
[CmdletBinding(SupportsShouldProcess=$true)]
param(
   [Parameter(
      Position=0,
      Mandatory=$true
   )]
   [String]$Path
   ,
   [Alias("query")]
   [Parameter(
      Position=1,
      Mandatory=$true
   )]
   [String]$SqlQuery="SELECT * FROM Sheet1"
   ,
   [Switch]$csv = $true
   ,
   [Alias("dlm")]
   [String]$delimiter=","

)
#endregion

# Error trap
trap
{
    # Restore culture
   Set-Culture $oldShortDatePattern $oldShortTimePattern $oldAMDesignator $oldPMDesignator
}

Function Set-Culture
{
   param(
     [string]$ShortDatePattern,
     [string]$ShortTimePattern,
     [string]$AMDesignator,
     [string]$PMDesignator
   )
   # save current settings
   [System.Globalization.DateTimeFormatInfo]$Culture=(Get-Culture).DateTimeFormat
   $script:oldShortDatePattern=$Culture.ShortDatePattern
   $script:oldShortTimePattern=$Culture.ShortTimePattern
   $script:oldAMDesignator=$Culture.AMDesignator
   $script:oldPMDesignator=$Culture.PMDesignator

   # set new settings
   $Culture.ShortDatePattern=$ShortDatePattern
   $Culture.ShortTimePattern=$ShortTimePattern
   $Culture.AMDesignator=$AMDesignator
   $Culture.PMDesignator=$PMDesignator
}

# Set Culture
Set-Culture "dd-MMM-yyyy" "HH:mm" "" ""

# Run Excel query
& {
$ErrorActionPreference = "Stop"

#$adOpenStatic = 3
#$adLockOptimistic = 3

# Run query
$SqlConnection = New-Object System.Data.OleDb.OleDbConnection

# Option 1
# $SqlConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=""$Path""; Extended Properties=""Excel 12.0; HDR=NO; IMEX=1; ReadOnly=True"" "

# Option 2
$SqlConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=""$Path""; Extended Properties=""Excel 12.0; HDR=YES; IMEX=1; ReadOnly=True"" "

$SqlCmd = New-Object System.Data.OleDb.OleDbCommand
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection

$SqlAdapter = New-Object System.Data.OleDb.OleDbDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd

$DataSet = New-Object System.Data.DataSet
$nRecs = $SqlAdapter.Fill($DataSet)
$nRecs | Out-Null

# Populate Hash Table
$objTable = $DataSet.Tables[0]

# Return results to console (pipe console output to Out-File cmdlet to create a file)
if ($csv) {
   # Option 1
   # ($objTable | ConvertTo-CSV -delimiter $delimiter -NoTypeInformation) -replace('"','') | Select-Object -skip 2

   # Option 2
   ($objTable | ConvertTo-CSV -delimiter $delimiter -NoTypeInformation) -replace('"','')
} else {
   # Option 1
   # $objTable | Select-Object -skip 2

   # Option 2
   $objTable
}

# Saves to File as CSV
# ($objTable | Export-CSV -Path $OutputPath -NoTypeInformation)  -replace('"','')

# Saves to File as XML
# $objTable | Export-Clixml -Path $OutputPath

} # end script block

# Restore culture
Set-Culture $oldShortDatePattern $oldShortTimePattern $oldAMDesignator $oldPMDesignator
