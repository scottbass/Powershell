Function SqlServerExecCommand
{
<# 
Header goes here...
#>
    [CmdletBinding( DefaultParameterSetName = 'Instance',
                    SupportsShouldProcess = $true,
                    ConfirmImpact = 'Medium' )]
    param(
        # Target
        [ValidateNotNullOrEmpty()]
        [Alias('ts','tgtsvr')]
        [String]
        $TgtServer              = 'SVDCMHPRRLSQD01'
        ,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias('tdb','tgtdb')]
        [String]
        $TgtDatabase
        ,
        [ValidateNotNullOrEmpty()]
        [Alias('tsc','tgtsch')]
        [String]
        $TgtSchema
        ,
        [ValidateNotNullOrEmpty()]
        [Alias('ttb','tgttbl')]
        [String]
        $TgtTable
        ,

        # Script options
        [Parameter(Mandatory=$true)]
        [Alias('sql','query')]
        [ValidateNotNullOrEmpty()]
        [String]
        $SqlQuery
        ,
        [Alias('timeout')]
        [int]
        $CommandTimeout         = 0 # unlimited
        ,
        [switch]
        $Quiet
    )

    ###############################################################################
    # Functions
    ###############################################################################
    Function ConnectionString([string] $ServerName, [string] $DbName)
    {
        "Data Source=$ServerName;Initial Catalog=$DbName;Integrated Security=True;Connection Timeout=120"
    }

    Function Print-Parms
    {
        if ($quiet) {return}
        if (! $verbose) {return}

        # list the parameters
        $name= @{Label='Parameter';  Expression={$_.Name};  Width=30}
        $value=@{Label='Value';      Expression={$_.Value}; Width=200}

        Push-Location variable:
        Get-Item `
            TgtServer,TgtDatabase,TgtTable,`
            CommandTimeout,`
            SqlQuery | `
            Format-Table $name,$value -Wrap | Out-Host
        Pop-Location
    }

    ###############################################################################
    # Initialization
    ###############################################################################
    Set-StrictMode -Version Latest
    $ErrorActionPreference='Stop'

    # Get verbose status
    $verbose = $VerbosePreference -ne 'SilentlyContinue'

    # Startup message
    if (! $quiet) {
        $msg = @'
[*] Script:    Started at {0}
'@ -f $(Get-Date)
        Write-Host ($msg) -BackgroundColor Cyan -ForegroundColor Black
    }

    # Resolve any embedded variables in $SqlQuery
    $SqlQuery = Invoke-Expression "Write-Output `"$SqlQuery`""

    # Print parameters
    Print-Parms

    if ($PSCmdlet.ShouldProcess(
        ('`n',$SqlQuery), 'Execute Command'
    ))
    {
        # Create source objects
        $TgtConn = New-Object System.Data.SqlClient.SQLConnection
        $TgtCmd  = New-Object System.Data.SqlClient.SqlCommand

        # Open connection to Sql Server
        $TgtConn.ConnectionString = ConnectionString $TgtServer $TgtDatabase
        $TgtCmd.Connection = $TgtConn
        $TgtConn.Open()
        
        # Split SQL query into batches based on the GO statement (case-insensitive split)
        $batches = $SqlQuery -split 'GO\r\n'

        # Execute each batch
        foreach($batch in $batches)
        {
            if ($batch.Trim() -ne ''){
                # Set command to execute
                $TgtCmd.CommandText = $batch
                $TgtCmd.CommandTimeout = $CommandTimeout

                # Execute command
                Try {
                    [Void]$TgtCmd.ExecuteNonQuery()
                }
                Catch [System.Data.SqlClient.SqlException]
                {
                    $ex = $_.Exception
                    if ($ex.Number -eq -2) {
                        $msg = @'
[*] Script:    Timeout executing command.
'@
                        Write-Host ($msg) -BackgroundColor Yellow -ForegroundColor Black
                    } else {
                        Throw $ex
                    }
                }
                Catch [System.Exception]
                {
                    $msg = @'
[*] Script:    Error executing command.
'@
                    Write-Host ($msg) -BackgroundColor Red -ForegroundColor Black

                    $ex = $_.Exception
                    Throw $ex
                }
            }
        }    

        # Close connections
        if (Test-Path variable:TgtConn)     {$TgtConn.Close();$TgtConn.Dispose()}
        if (Test-Path variable:TgtCmd)      {$TgtCmd.Dispose()}
        [System.GC]::Collect()

        if (! $quiet) {
            $msg = @'
[*] Script:    Ended   at {0}
'@ -f $(Get-Date)
            Write-Host ($msg) -BackgroundColor Green -ForegroundColor Black
        }
    }         
}

### END OF FILE ###
