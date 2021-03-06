<# HEADER
/*=====================================================================
Program Name            : Query-SQLServer.ps1
Purpose                 : Execute query against SQL Server
Powershell Version:     : v2.0
Input Data              : N/A
Output Data             : N/A

Originally Written by   : Scott Bass
Date                    : 15APR2013
Program Version #       : 1.0

=======================================================================

Modification History    :

=====================================================================*/
#>

<#
.SYNOPSIS
Query SQL Server

.DESCRIPTION
Execute a query against SQL Server

.PARAMETER  SQLQuery
SQL Query to execute

.PARAMETER  SQLServer
SQL Server Instance

.PARAMETER  SQLDatabase
SQL Database Instance

.PARAMETER  Csv
Output as CSV?  If no, the Dataset Table object is returned to the pipeline

.PARAMETER  Whatif
Echos the SQL query information without actually executing it.

.PARAMETER  Confirm
Asks for confirmation before actually executing the query.

.PARAMETER  Verbose
Prints the SQL query to the console window as it executes it.

.EXAMPLE
.\Query-SQLServer.ps1

Description
-----------
Queries the default SQL Server with the default query.

.EXAMPLE
.\Query-SQLServer.ps1 "select * from sys.all_tables"

Description
-----------
Queries the default SQL Server with the specified query.

.EXAMPLE
.\Query-SQLServer.ps1 "select * from dbo.jobs" -server AUMELBCASAS01 -database JAMS

Description
-----------
Queries the specified SQL Server instance and database with the specified query.

.EXAMPLE
.\Query-SQLServer.ps1 -csv:$false

Description
-----------
Queries the default SQL Server with the default query, 
returning the Object Table to the pipeline.

#>

#region Parameters
[CmdletBinding(SupportsShouldProcess=$true)]
param(
   [Alias("query")]
   [Parameter(
      Position=0
   )]
   [String]$SqlQuery = "SELECT * FROM dbo.Hist"
   ,
   [Alias("server")]
   [Parameter(
      Position=1
   )]
   [String]$SqlServer = "AUMELBCASAS02"
   ,
   [Alias("database")]
   [Parameter(
      Position=2
   )]
   [String]$SqlDatabase = "JAMS"
   ,
   [Switch]$csv = $true
)
#endregion

# Example of a generated query from the SQL Server client, pasted into this script
# Uses a here-document (@" ... "@)
<#
$SqlQuery = @"
SELECT [job_name]
      ,[restart_cnt]
      ,[jams_entry]
      ,[master_entry]
      ,[sched_time]
      ,[hold_time]
      ,[start_time]
      ,[completion_time]
      ,[master_ron]
      ,[ron]
      ,[job_id]
      ,[setup_id]
      ,[cputim]
      ,[pageflts]
      ,[wspeak]
      ,[biocnt]
      ,[diocnt]
      ,[finalsts]
      ,[finalsev]
      ,[pid]
      ,[submitted_by]
      ,[username]
      ,[batch_que]
      ,[debug]
      ,[nodename]
      ,[log_filename]
      ,[ovr_job_name]
      ,replace(replace(finalsts_text,char(10),''),char(13),'') as [finalsts_text]
      ,[note_text]
      ,[job_status]
      ,[jams_id]
      ,[initiator_type]
      ,[initiator_id]
  FROM [JAMS].[dbo].[Hist]
"@
#>

$ErrorActionPreference = "Stop"

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Data Source=$SqlServer; Initial Catalog=$SqlDatabase; Integrated Security=True"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$nRecs = $SqlAdapter.Fill($DataSet)
$nRecs | Out-Null

# Populate Hash Table
$objTable = $DataSet.Tables[0]

# Return results to console (pipe console output to Out-File cmdlet to create a file)
if ($csv) {
   ($objTable | ConvertTo-CSV -NoTypeInformation) -replace('"','')
} else {
   $objTable
}
