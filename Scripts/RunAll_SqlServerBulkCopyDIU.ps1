<# 
Header goes here...
#>

[CmdletBinding( SupportsShouldProcess = $true,
                ConfirmImpact = 'Medium' )]

param(
    # Tables to process
    [System.Array]
    $Tables
    ,

    # Source
    [ValidateNotNullOrEmpty()]
    [Alias('ss','srcsvr')]
    [String]
    $SrcServer              = 'DOHNSCLDBSASBI,54491'
    ,
    [ValidateNotNullOrEmpty()]
    [Alias('sdb','srcdb')]
    [String]
    $SrcDatabase            = 'HIERep_prod' 
    ,
    [ValidateNotNullOrEmpty()]
    [Alias('ssc','srcsch')]
    [String]
    $SrcSchema              = 'dbo'
    ,

    # Target
    [ValidateNotNullOrEmpty()]
    [Alias('ts','tgtsvr')]
    [String]
    $TgtServer              = 'SVDCMHPRRLSQD01'
    ,
    [ValidateNotNullOrEmpty()]
    [Alias('tdb','tgtdb')]
    [String]
    $TgtDatabase            = 'RLCS_dev'
    ,
    [ValidateNotNullOrEmpty()]
    [Alias('tsc','tgtsch')]
    [String]
    $TgtSchema              = 'content' 
    ,

    # Script options
    [Alias('timeout')]
    [int]
    $CommandTimeout         = 120 # mainly used when retrieving initial row count
                                  # set to a smaller value if you do not want to wait as long
    ,
    [switch]
    $Clone                  = $true 
    ,
    [switch]
    $Rename                 = $false 
    ,
    [switch]
    $Truncate               = $false
    ,
    [switch]
    $Quiet                  = $false
    ,

    # SqlBulkCopy constructor options
    # See https://msdn.microsoft.com/en-us/library/system.data.sqlclient.sqlbulkcopyoptions(v=vs.110).aspx
    [switch]
    $CheckConstraints       = $true
    ,
    [switch]
    $FireTriggers           = $false
    ,
    [switch]
    $KeepIdentity           = $true
    ,
    [switch]
    $KeepNulls              = $true
    ,
    [switch]
    $TableLock              = $true
    ,

    # SqlBulkCopy runtime properties
    # See https://msdn.microsoft.com/en-us/library/system.data.sqlclient.sqlbulkcopy(v=vs.110).aspx
    [int]
    $BatchSize              = 1000000   # all data written in one transaction
    ,
    [int]
    $NotifyAfter            = 1000000   # progress report every 1M records processed
                                        # 0 = no progress reports, including the progress bar
                                        # -verbose must be specified for the console output
    ,
    [int]
    $BulkCopyTimeout        = 18000     # 5 hours
)

# Source the parameters
# (this may override some command line options)
. \\sascs\linkage\RL_content_snapshots\Powershell\Scripts\RunAll_SqlServerParametersDIU.ps1

# Source the SqlServer function
. \\sascs\linkage\RL_content_snapshots\Powershell\Functions\SqlServerBulkCopy.ps1

# Continue this script if an error occurs
$ErrorActionPreference = 'Continue'

# Hardcode verbose output
$VerbosePreference = 'Continue'

# Inline Functions
Function RunSqlServerBulkCopy 
{
    # Function to run SqlServerBulkCopy
    [CmdletBinding( SupportsShouldProcess = $true,
                    ConfirmImpact = 'Medium' )]

    param(
        # Source
        [ValidateNotNullOrEmpty()]
        [String]$table
    )

    # Lookup the custom object
    $obj=$ht.Get_Item($table)

    $parms=@()

    $WhatIf  = $WhatIfPreference.IsPresent
    $Verbose = $VerbosePreference -ne 'SilentlyContinue'

    # Should Process parameters
    if ($WhatIf)                 {$parms+='-WhatIf'}
    if ($Verbose)                {$parms+='-Verbose'}

    # Switch parameters
    if ($obj.Clone)              {$parms+='-Clone'}
    if ($obj.Rename)             {$parms+='-Rename'}
    if ($obj.Truncate)           {$parms+='-Truncate'}
    if ($obj.Quiet)              {$parms+='-Quiet'}

    if ($obj.CheckConstraints)   {$parms+='-CheckConstraints'}
    if ($obj.FireTriggers)       {$parms+='-FireTriggers'}
    if ($obj.KeepIdentity)       {$parms+='-KeepIdentity'}
    if ($obj.KeepNulls)          {$parms+='-KeepNulls'}
    if ($obj.TableLock)          {$parms+='-TableLock'}

    # Named parameters
    # I know these all have values in the custom object so the if statement is true,
    # I'm just trying to future-proof this code against future changes
    if ($obj.BatchSize)          {$parms+='-BatchSize';        $parms+=$obj.BatchSize}
    if ($obj.NotifyAfter)        {$parms+='-NotifyAfter';      $parms+=$obj.NotifyAfter}
    if ($obj.CommandTimeout)     {$parms+='-CommandTimeout';   $parms+=$obj.CommandTimeout}
    if ($obj.BulkCopyTimeout)    {$parms+='-BulkCopyTimeout';  $parms+=$obj.BulkCopyTimeout}

    if ($obj.SrcServer)          {$parms+='-SrcServer';        $parms+="""{0}""" -f $obj.SrcServer}  # embedded comma
    if ($obj.SrcDatabase)        {$parms+='-SrcDatabase';      $parms+=$obj.SrcDatabase}
    if ($obj.SrcSchema)          {$parms+='-SrcSchema';        $parms+=$obj.SrcSchema}
    if (! $obj.SrcTable)         {$obj.SrcTable=$table}
                                  $parms+='-SrcTable';         $parms+=$obj.SrcTable

    if ($obj.TgtServer)          {$parms+='-TgtServer';        $parms+="""{0}""" -f $obj.TgtServer}  # embedded comma
    if ($obj.TgtDatabase)        {$parms+='-TgtDatabase';      $parms+=$obj.TgtDatabase}
    if ($obj.TgtSchema)          {$parms+='-TgtSchema';        $parms+=$obj.TgtSchema}   
    if (! $obj.TgtTable)         {$obj.TgtTable=$table}
                                  $parms+='-TgtTable';         $parms+=$obj.TgtTable

    # Get SQL Query from external file
    $OFS='`r`n'  # preserve embedded CRLF's
    $local:SqlQuery = Get-Content "\\sascs\linkage\RL_content_snapshots\SQLServer\RLCS\SqlBulkCopy_${table}.sql"
    if ($obj.TgtServer -eq $obj.SrcServer) {
       $local:SqlQuery = ($local:SqlQuery -f '', $obj.SrcDatabase, $obj.SrcSchema, $obj.SrcTable) -replace '\[\]\.',''
    } else {
       $local:SqlQuery = ($local:SqlQuery -f $obj.SrcServer, $obj.SrcDatabase, $obj.SrcSchema, $obj.SrcTable) -replace '\[\]\.',''
    }
    $OFS=$null
                                  $parms+='-SqlQuery';         $parms+="""{0}""" -f $SqlQuery  # embedded single quotes
                                  
    $s='{1}.{2}.{3}' -f $obj.SrcServer, $obj.SrcDatabase, $obj.SrcSchema, $obj.SrcTable
    $t='{1}.{2}.{3}' -f $obj.TgtServer, $obj.TgtDatabase, $obj.TgtSchema, $obj.TgtTable
    $msg='From {0} to {1}' -f $s,$t

    if ($PSCmdlet.ShouldProcess($msg,'SQL Bulk Copy'))
    {
        Try {
            Invoke-Expression "SqlServerBulkCopy $parms"
        }
        Catch {
            $msg = $_.Exception.Message
            Write-Error "Error copying $s to $t"
            Write-Error ($msg)
            $ErrorActionPreference = 'Continue'
            # Throw $_
        }
        Finally {
            return
        }
    }
}

###############################################################################
# MAIN PROCESSING
###############################################################################
foreach ($table in $Tables) {
    $table = $table.ToUpper()
    
    # If the table is not defined print warning and return
    if ($ht.Get_Item($table) -eq $null) {
        Write-Warning "Table $table is not defined in the metadata.  Skipping..."
    } else {
        RunSQLServerBulkCopy($table)
    }
}

### END OF FILE ###
