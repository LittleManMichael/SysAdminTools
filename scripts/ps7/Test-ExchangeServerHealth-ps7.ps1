function Test-ExchangeServerHealth-ps7
{
    [cmdletbinding()]
    param (
        [parameter()][string]$Server,
        [parameter()][string]$ServerList,
        [parameter()][string]$SendEmail
    )

    # Server/Serverlist both used check
    if ($Server -and $ServerList)
    {
        throw "The parameters '-Server' and '-ServerList' cannot be used together. Please use one or the other."
    }

    Write-Host "Initializing..."

    # Gather ESN Mailboxes to use as remote connection(s)
    $esnMBX   = [System.Collections.Generic.List[System.String]]@()
    $mbxRelay = @(((Resolve-DnsName -name relay.newton.pentagon.mil).IPAddress | foreach { Resolve-DnsName -name $_ }).NameHost)

    # Remove any previous remote sessions
    $previousSessions = Get-PSSession | Where ConfigurationName -eq "Microsoft.Exchange"
    if ($previousSessions) { $previousSessions | Remove-PSSession }

    # Attempt to gather remote commands from MBX server(s)
    $remoteCommands = $false
    for ($i = 0; $remoteCommands -eq $false; $i++)
    {
        try
        {
            $psSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($mbxRelay[$i])/powershell"
            $import    = Import-PSSession $psSession -InformationAction Ignore -WarningAction Ignore -AllowClobber
            Write-Host "Connected to $($mbxRelay[$i]) for remote exchange commands." -f Green
            $remoteCommands = $true
        }

        catch 
        {
            Write-Warning "Unable to utilize $($mbxRelay[$i]) for remote exchange commands.. Will try another server."
            if ($i -ge $mbxRelay) { throw "Ran out of servers to try... Please check status of MBX Servers." }
            $remoteCommands = $false
        }
    }
    
    #############
    # VARIABLES #
    #############
    $user    = $env:USERNAME -replace '-admin',''
    $now     = Get-Date
    $pass    = "Green"
    $warn    = "Yellow"
    $fail    = "Red"
    $success = "Success"
    $ip      = $null

    $myDir          = "\\newton\admin\SysAdmin Powershell Module\Scripts\Test-ExchangeServerHealth-ps7"
    $ignoreListFile = Join-Path $myDir 'ignorelist.txt'
    $ignoreList     = @(Get-Content -Path $ignoreListFile -ErrorAction SilentlyContinue)

    [int]$mapiTimeOut = 10
    [int]$qHigh       = 100
    [int]$qWarn       = 80
    [int]$replayQwarn = 8

    $exchangeServers = [System.Collections.Generic.List[System.Object]]@()
    $serverSummary   = [System.Collections.Concurrent.ConcurrentBag[object]]::New()
    $serverReport    = [System.Collections.Concurrent.ConcurrentBag[object]]::New()
    $dagSummary      = [System.Collections.Concurrent.ConcurrentBag[object]]::New()
    $dagDatabase     = [System.Collections.Concurrent.ConcurrentBag[object]]::New()
    $dagDBCopyReport = [System.Collections.Concurrent.ConcurrentBag[object]]::New()
    $dagCiReport     = [System.Collections.Concurrent.ConcurrentBag[object]]::New()
    $dagMemberReport = [System.Collections.Concurrent.ConcurrentBag[object]]::New()

    $casRole = "Client Access Server Role"
    $htRole  = "Hub Transport Server Role"
    $mbRole  = "Mailbox Server Role"
    $umRole  = "Unified Messaging Server Role"

    $smtpsettings = @{
	    To         = "msprous@newton.pentagon.mil" #"-nocsustainment@newton.pentagon.mil"
	    From       = "-nocsustainment@newton.pentagon.mil"
	    Subject    = "Exchange Server Health Report - $now"
	    SmtpServer = "relay.newton.pentagon.mil"
	}
    
    #############
    # FUNCTIONS #
    #############
    function New-ServerHTMLTableCell
    {
        param ([string]$result, [switch]$isINT)

        # Reset Variable
        $htmlTableCell = $null 

        # Check if INT
        if ($isINT)
        {
            # Switch based off number
            switch ($result -as [int])
            {
                {$_ -ge 24} { $htmlTableCell = "<td class=""pass"">$($result)</td>" } # Pass
                {$_ -lt 24} { $htmlTableCell = "<td class=""warn"">$($result)</td>" } # Warn
                
                default     { $htmlTableCell = "<td>$($result)</td>" } # if none of the above
            }

            return $htmlTableCell
        }

        # Default logic
        switch ($result)
        {
            {$_ -like "PASS*"} { $htmlTableCell = "<td class=""pass"">$($result)</td>" } # Pass
            {$_ -like "WARN*"} { $htmlTableCell = "<td class=""warn"">$($result)</td>" } # Warn
            {$_ -like "FAIL*"} { $htmlTableCell = "<td class=""fail"">$($result)</td>" } # Fail

            "SUCCESS"          { $htmlTableCell = "<td class=""pass"">$($result)</td>" } # Pass
            "N/A"              { $htmlTableCell = "<td class=""warn"">$($result)</td>" } # Warn
            "Access Denied"    { $htmlTableCell = "<td class=""fail"">$($result)</td>" } # Fail

            default            { $htmlTableCell = "<td>$($result)</td>" } # if none of the above
        }

        return $htmlTableCell
    }

    function New-DagHTMLTableCell
    {
        param ($result, [switch]$healthy, [switch]$isINT, [switch]$Summary, [switch]$Details, [switch]$Members)

        # Reset Varaible
        $htmlTableCell = $null
        $replayQwarn = 8

        # DAG Summary Table
        if ($Summary)
        {
            switch($result[0])
            {
                {($_ -eq $result[-1]) -and ($healthy)}        { $htmlTableCell = "<td class=""pass"">$($result[0])</td>" } # Pass
                {($_ -eq 0) -and ($healthy)}                  { $htmlTableCell = "<td class=""fail"">$($result[0])</td>" } # Fail

                {($_ -eq 0) -and (-not ($healthy))}           { $htmlTableCell = "<td class=""pass"">$($result[0])</td>" } # Pass
                {($_ -eq $result[-1]) -and (-not ($healthy))} { $htmlTableCell = "<td class=""fail"">$($result[0])</td>" } # Fail

                default                                       { $htmlTableCell = "<td class=""warn"">$($result[0])</td>" } # if none of the above
            }

            return $htmlTableCell
        }

        # DAG Details Table
        if ($Details)
        {
            switch ($result)
            {
                {$_ -eq 'Mounted' -or $_ -eq 'Healthy'}  { $htmlTableCell = "<td class=""pass"">$($result)</td>" } # Pass
                {$_ -ne 'Mounted' -and $_ -ne 'Healthy'} { $htmlTableCell = "<td class=""fail"">$($result)</td>" } # Fail

                {$isINT -and $_ -lt $replayQwarn}        { $htmlTableCell = "<td class=""pass"">$($result)</td>" } # Pass

                default                                  { $htmlTableCell = "<td class=""warn"">$($result)</td>" } # if none of the above
            }

            return $htmlTableCell
        }

        # DAG Members Table
        if ($Members)
        {
            switch ($result)
            {
                "Passed" { $htmlTableCell = "<td class=""pass"">$($result)</td>" } # Pass
                "N/A"    { $htmlTableCell = "<td class=""info"">$($result)</td>" } # N/A

                default  { $htmlTableCell = "<td class=""warn"">$($result)</td>" } # if none of the above
            }

            return $htmlTableCell
        }
    }

    ######################
    # SERVER COUNT CHECK #
    ######################
    # Single server
    if ($server)
    {
        try
        {
            $exchangeServers = @(Get-ExchangeServer -Identity $Server -ErrorAction Stop | Select-Object -ExpandProperty Name)
        }
        catch
        {
            throw "Error gathering $Server : $($_.Exception.Message)" 
        }
    }

    # ServerList
    elseif ($ServerList)
    {
        # Check for comma. If comma found, it's likely multiple separated by commands. otherwise, it's likely a file
        if ($ServerList -match ',')
        {
            $exchangeServers = @($ServerList -split ',')
        }
        else 
        {
            try 
            {
                $exchangeServers = @(Get-Content -Path $ServerList -ErrorAction Stop)
            }
            catch 
            {
                throw "Server List file '$ServerList' cannot be found or is invalid."
            }
        }
    }

    # Default
    else 
    {
        $tmpServers = @(Get-ExchangeServer | Sort-Object Site,Name)
        if ($ignoreList.Count -gt 0)
        {
            foreach ($svr in $tmpServers)
            {
                if ($ignoreList -notcontains $svr.Name)
                {
                    $exchangeServers.Add($svr.Name)
                }
            }
        }
        else
        {
            $exchangeServers = @(Get-ExchangeServer | Sort-Object Site,Name | Select-Object -ExpandProperty Name)
        }
    }

    cls # Clear the script

    #######################################
    # Begin Parallel Server Health Checks #
    #######################################
    Write-Host "RUNNING EXCHANGE SERVER CHECKS" -F Yellow
    Write-Host "------------------------------"

    $exchangeServers | ForEach-Object -Parallel `
    {
        $server = $_ # Assign $server variable for easier reading

        # Pull collections into current runspace
        $serverReportBag  = $using:serverReport
        $serverSummaryBag = $using:ServerSummary

        # Gain Remote session and get server information
        $psSession  = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($server)/powershell" -Authentication Kerberos
        $serverInfo = Invoke-Command -Session $psSession -ArgumentList $server { param($svr) Get-ExchangeServer -Identity $svr }

        # Create the initial Server Object
        $serverObj = [pscustomobject] `
        @{ 
            'Server'          = $server.ToUpper()
            'Site'            = if ($serverInfo) { ($serverInfo.Site.ToString() -split '/')[-1] } else { "N/A" }
            'DNS'             = $null
            'Ping'            = $null
            'Uptime (hrs)'    = $null
            'CASR Services'   = "N/A"
            'HTSR Services'   = "N/A"
            'MBSR Services'   = "N/A"
            'Transport Queue' = "N/A"
            'MB DBs Mounted'  = "N/A"
            'Mail Flow Test'  = "N/A"
            'MAPI Test'       = "N/A"
        }

        # Gather DNS property
        try
        {
            $ip = [System.Net.Dns]::GetHostAddresses($server) | ForEach-Object { $_.IPAddressToString }
            if (-not $ip) 
            {
                $serverObj.DNS = 'FAIL'
                $serverReportBag.Add($serverObj)
                $serverSummaryBag.Add("$server - Server not found in DNS.")
            }
            $serverObj.DNS = 'PASS'
        }
        catch
        {
            $serverObj.DNS = 'FAIL'
            $serverReportBag.Add($serverObj)
            $serverSummaryBag.Add("$server - Server not found in DNS.")
            break
        }

        # Ping Test
        $PingTest = Test-Connection -ComputerName $server -count 3 -Quiet -ErrorAction SilentlyContinue
        $serverObj.Ping = if ($PingTest) { 'PASS' } else { 'FAIL' ; $serverReportBag.Add($serverObj) ; $serverSummaryBag.Add("$server - Ping Failed.") ; break }

        # Begin Server Uptime Check
        try
        {
            $OS = Get-CimInstance -ComputerName $server -ClassName Win32_OperatingSystem -ErrorAction Stop
            $timeSpan = (Get-Date) - $OS.LastBootUpTime
            [int]$serverObj.'Uptime (hrs)' = [math]::Floor($timeSpan.TotalHours) 
        }
        catch
        {
            $serverObj.'Uptime (hrs)' = 'FAIL'
            $serverSummaryBag.Add("$server - Failed to retrieve uptime")
        }

        # Test Services health
        try
        {
            $serviceInfo = Invoke-Command -Session $psSession -ArgumentList $server { param($svr) Test-ServiceHealth -Server $svr -ErrorAction STOP }

            # Gather info
            $casr = $serviceInfo | Where-Object { $_.Role -eq 'Client Access Server Role' }
            $htsr = $serviceInfo | Where-Object { $_.Role -eq 'Hub Transport Server Role' }
            $mbsr = $serviceInfo | Where-Object { $_.Role -eq 'Mailbox Server Role' }

            # Check Failed Services
            # Client Access Server Role
            switch ($casr)
            {
                {$_.ServicesNotRunning.Count -eq 0} { $serverObj.'CASR Services' = "PASS" }
                
                default { $_.ServicesNotRunning | 
                    ForEach-Object { $serverSummaryBag.Add("$server - '$_' service is not running") } ; $serverObj.'CASR Services' = "FAIL" }
            } 

            # Hub Transport Server Role
            switch ($htsr)
            {
                {$_.ServicesNotRunning.Count -eq 0} { $serverObj.'HTSR Services' = "PASS" }
                
                default { $_.ServicesNotRunning | 
                    ForEach-Object { $serverSummaryBag.Add("$server - '$_' service is not running") } ; $serverObj.'HTSR Services' = "FAIL" }
            }

            # Mailbox Server Role
            switch ($mbsr)
            {
                {$_.ServicesNotRunning.Count -eq 0} { $serverObj.'MBSR Services' = "PASS" }
                
                default { $_.ServicesNotRunning | 
                    ForEach-Object { $serverSummaryBag.Add("$server - '$_' service is not running") } ; $serverObj.'MBSR Services' = "FAIL" }
            }
        }
        catch 
        {
            $serverObj.'CASR Services' = "FAIL"
            $serverobj.'HTSR Services' = "FAIL"
            $serverobj.'MBSR Services' = "FAIL"
            $serverSummaryBag.Add("$server - Failed to get service info.")
        }


        # Check Hub Transport
        try 
        {
            # Get the queues
            $queue = Invoke-Command -Session $psSession -ArgumentList $server { param($svr) Get-Queue -Server $svr -ErrorAction STOP }
        
            # Get total count
            [int]$qcnt = ($queue | Measure-Object MessageCount -Sum).Sum

            # Define thresholds
            [int]$qHigh = 100
            [int]$qWarn = 80

            # Define answer based on count
            switch ($qcnt)
            {
                { $_ -lt $qWarn }                    { $serverObj.'Transport Queue' = "PASS ($qcnt)" }
                { $_ -ge $qHigh }                    { $serverObj.'Transport Queue' = "FAIL ($qcnt)" ; $serverSummaryBag.Add("$server - Transport Queue is severly above warning threshold") }
                { $_ -ge $qWarn -and $_ -lt $qHigh } { $serverObj.'Transport Queue' = "WARN ($qcnt)" ; $serverSummaryBag.Add("$server - Transport Queue is above warning threshold") }
                default                              { $serverObj.'Transport Queue' = "FAIL" ; $serverSummaryBag.Add("$server - Failed to gather Transport Queue") }
            }
        }
        catch
        {
            $serverObj.'Transport Queue' = "Unknown"
            $serverSummaryBag.Add("$server - Failed to gather Transport Queue")
        }

        # Check Mailboxes
        # Get PF and MB Databases
        [array]$pfDBs = @(Invoke-Command -Session $psSession -ArgumentList $server { param($svr) Get-PublicFolderDatabase -Server $svr -Status -WarningAction SilentlyContinue })
        [array]$mbDBs = @(Invoke-Command -Session $psSession -ArgumentList $server { param($svr) Get-MailboxDatabase -Server $svr -Status }) | Where-Object { $_.Recovery -ne $true }
        [array]$activeDBs = @(Invoke-Command -Session $psSession -ArgumentList $server,$serverInfo { param($svr,$serverInfo) Get-MailboxDatabase -Server $svr -Status }) | Where-Object { $_.Recovery -ne $true -and $_.MountedOnServer -eq $serverInfo.fqdn }

        # Check the mailbox databases
        if ($mbDBs.Count -gt 0)
        {
            [string]$mbdbStatus = "PASS"
            $alertDBs = [system.Collections.Generic.List[string]]::New()

            foreach ($db in $mbDBs)
            {
                if ($db.mounted -ne $true)
                {
                    $mbdbStatus = "FAIL"
                    $alertDBs.Add($db.Name)
                    $serverSummaryBag.Add("$server - Mailbox Database '$($db.Name)' not mounted")
                }
            }

            $serverObj.'MB DBs Mounted' = $mbdbStatus
        }

        # MAPI Connectivity Test
        if ($activeDBs.Count -gt 0 -or $pfDBs.Count -gt 0)
        {
            foreach ($db in $mbDBs)
            {
                # Gather MAPI Status
                $mapiStatus = Invoke-Command -Session $psSession -ArgumentList $db.Identity { param($db) Test-MapiConnectivity -Database $db -PerConnectionTimeout 10 }
                
                # Go through Status Result
                if ($mapiStatus.Result.Value -eq $null)
                {
                    $mapiResult = $mapiStatus.Result
                }
                else 
                {
                    $mapiResult = $mapiStatus.Result.Value
                }

                if (($mapiResult) -ne "Success")
                {
                    $mapiStatus = "FAIL"
                    $alertDBs.Add($db.Name)
                    $serverSummaryBag.Add("$server - Mailbox Database '$($db.Name)' failed MAPI Connectivity")
                }
                else
                {
                    $mapiStatus = "PASS"
                }

                # Add the property to serverobj
                $serverObj.'MAPI Test' = $mapiStatus
            }
        }

        # Mail Flow Test
        if ($activeDBs.Count -gt 0)
        {
            try
            {
                $mailFlow = Invoke-Command -Session $psSession -ArgumentList $server { param($svr) Test-Mailflow -Identity $svr -ErrorAction STOP }
                switch ($mailFlow.TestMailFlowResult)
                {
                    'Success' { $serverObj.'Mail Flow Test' = "PASS" }
                    default   { $serverObj.'Mail Flow Test' = "FAIL" ; $serverSummaryBag.Add("$server - Mail Flow Test Failed.") }
                }
            }
            catch
            {
                $serverObj.'Mail Flow Test' = "FAIL"
                $serverSummaryBag.Add("$server - Mail Flow Test Failed.")
            }
        }

        Write-Host "$($server.ToUpper()) - CHECKS COMPLETE" -f Green
        $serverReportBag.Add($serverObj)
    } -ThrottleLimit 30

    Write-Host "`nEXCHANGE SERVER CHECKS COMPLETED" -f Green
    Write-Host "__________________________________`n`n"

    ########################
    # End of Server Checks #
    ########################

    ####################
    # Begin DAG Checks #
    ####################
    
    Write-Host "RUNNING DAG CHECKS" -f Yellow
    Write-Host "------------------"

    # Get all DAGs and remove ignored DAGs
    $tmpDags = @(Invoke-Command -Session $psSession { Get-DatabaseAvailabilityGroup })
    $dags    = $tmpDags | Where-Object { (!($ignoreList -icontains $_.Name)) }

    # Get all mailboxes and remove ignored databases
    $tmpDatabases = @(Invoke-Command -Session $psSession { Get-MailboxDatabase -Status -IncludePreExchange2013 })
    $dagDatabases = [System.Collections.Concurrent.ConcurrentBag[object]]@($tmpDatabases | Where-Object { (!($ignorelist -icontains $_.Name)) })

    # Check DAG count. If 0, output so. Otherwise being parallel checks
    if ($dags.Count -eq 0)
    {
        Write-Host "No DAGs to check..." -f Gray
    }

    if ($dags.Count -gt 0)
    {
        # Being Parallel DAG Checks
        $dags | ForEach-Object -Parallel `
        {
            # Setup variables and pull collections
            $dag                = $_
            $replayQwarn   = $using:replayQwarn 
            $dagServer          = $dag.Servers[0] 
            $dagDatabasesBag    = $using:dagDatabases
            $dagSummaryBag      = $using:dagSummary 
            $dagDBcopyReportBag = $using:dagDBcopyReport
            $dagCiReportBag     = $using:dagCIReport
            $dagMemberReportBag = $using:dagMemberReport

            # Connect to remote commands
            $psSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$dagServer/powershell" -Authentication Kerberos

            # Get all Databases in the DAG and DAG Members
            $dagDBs     = @($dagDatabasesBag | Where-Object { $_.MasterServerOrAvailabilityGroup -eq $dag.Name } | Sort Name).Name 
            $dagMembers = @($dag | Select-Object -ExpandProperty Servers | Sort-Object Name)

            # Go through each DB in the dag
            foreach ($db in $dagDBs)
            {
                # Create the Object
                $dbObj = [pscustomobject] `
                @{
                    'Database'         = $db 
                    'Mounted On'       = "Unknown"
                    'Total copies'     = $null
                    'Healthy Copies'   = $null
                    'Unhealthy Copies' = $null
                    'Healthy Queues'   = $null
                    'Unhealthy Queues' = $null
                    'DAG'              = $dag.Name 
                    'Copy Queue'       = $null    
                }

                # Gather DB copy status'
                $dbCopyStatus = @(Invoke-Command -Session $psSession -ArgumentList $db { param($dbname) Get-MailboxDatabaseCopyStatus -Identity $dbname })

                # Go through each db copy
                foreach ($dbCopy in $dbCopyStatus)
                {
                    # Create the dbCopy obj
                    $dbCopyObj = [pscustomobject] `
                    @{
                        'Database Copy'  = $dbCopy.Identity
                        'Database Name'  = $dbCopy.DatabaseName
                        'Mailbox Server' = $dbCopy.MailboxServer
                        'Status'         = $dbCopy.Status 
                        'Copy Queue'     = [int]$dbCopy.CopyQueueLength
                        'Replay Queue'   = [int]$dbCopy.ReplayQueueLength
                        'DAG'            = $dag.Name  
                    }

                    # Add DAG to summary if issues found below
                    # Queue Warnings
                    switch ($dbCopyObj) 
                    {
                        {$_.'Copy Queue' -gt $replayQwarn }   { $dagSummaryBag.Add("$($dbCopyObj.'Database Copy') - Copy Queue is at $($dbCopyObj.'Copy Queue')") }
                        {$_.'Replay Queue' -gt $replayQwarn } { $dagSummaryBag.Add("$($dbCopyObj.'Database Copy') - Replay Queue is at $($dbCopyObj.'Replay Queue')") }
                    }

                    # Status Warnings
                    if ($dbCopyObj.Status -ne 'Mounted' -and $dbCopyObj.Status -ne 'Healthy') 
                    {
                        $dagSummaryBag.Add("$($dbCopyObj.'Database Copy') - Status is $($dbCopyObj.Status)")
                    }

                    # Add the dbcopyobj to the report
                    $dagDBcopyReportBag.Add($dbCopyObj)
                }

                # Gather all the copies for the current database
                $copies = @($dagDBcopyReportBag | Where-Object { $_.'Database Name' -eq $db })

                # Add the final dbObj properties
                $dbObj.'Mounted On'       = ($copies | Where-Object { $_.Status -eq 'Mounted'}).'Mailbox Server'
                $dbObj.'Total copies'     = $copies.Count
                $dbObj.'Healthy Copies'   = @($copies | Where-Object { ($_.Status -eq 'Mounted') -or ($_.Status -eq 'Healthy') }).Count 
                $dbObj.'Unhealthy Copies' = @($copies | Where-Object { ($_.Status -ne 'Mounted') -and ($_.Status -ne 'Healthy') }).Count 
                $dbObj.'Healthy Queues'   = @($copies | Where-Object { ($_.'Copy Queue' -lt $replayQwarn) -and ($_.'Replay Queue' -lt $replayQwarn) }).Count
                $dbObj.'Unhealthy Queues' = @($copies | Where-Object { ($_.'Copy Queue' -ge $replayQwarn) -and ($_.'Replay Queue' -ge $replayQwarn) }).Count

                # Add to DAG Summary if issues below found
                switch ($dbObj)
                {
                    # Copies
                    { $_.'Healthy Copies' -lt $_.'Total Copies' } { $dagSummaryBag.Add("$($dbObj.Database) - Healthy Copy count is $($dbObj.'Healthy Copies') (of $($dbObj.'Total copies'))") }
                    
                    # Queues
                    { $_.'Healthy Queues' -lt $_.'Total Copies' } { $dagSummaryBag.Add("$($dbObj.Database) - Healthy Queue count is $($dbObj.'Healthy Queues') (of $($dbObj.'Total copies'))") }
                }

                # Add the DBobj to the report
                $dagDatabasesBag.Add($dbObj)
            }

            # Go through each DAG Member
            foreach ($dagMem in $dagMembers)
            {
                # Reset variable
                $repHealth = $null

                # Run the replication health test
                $repHealth = Invoke-Command -Session $psSession -ArgumentList $dagMem { param($dagmem) Test-ReplicationHealth -Identity $dagmem }

                # Create and populate object
                $memberObj = [pscustomobject] `
                @{
                    Server = $dagMem 

                    # if statements below are because some servers don't show these results, so we have to check for noshows
                    ClusterService    = if (-not ($repHealth.Check -contains 'ClusterService')) { "Unknown" }
                                        else { ($repHealth | Where-Object { $_.Check -eq 'ClusterService' }).Result.ToUpper() }
                    
                    ReplayService     = if (-not ($repHealth.Check -contains 'ReplayService')) { "Unknown" }
                                        else { ($repHealth | Where-Object { $_.Check -eq 'ClusterService' }).Result.ToUpper() }
                    
                    ActiveManager     = if (-not ($repHealth.Check -contains 'ActiveManager')) { "Unknown" }     
                                        else { ($repHealth | Where-Object { $_.Check -eq 'ActiveManager' }).Result.ToUpper() }
                    
                    TasksRpcListener     = if (-not ($repHealth.Check -contains 'TasksRpcListener')) { "Unknown" }
                                           else { ($repHealth | Where-Object { $_.Check -eq 'TasksRpcListener' }).Result.ToUpper() }
                    
                    TcpListener     = if (-not ($repHealth.Check -contains 'TcpListener')) { "Unknown" }
                                      else { ($repHealth | Where-Object { $_.Check -eq 'TcpListener' }).Result.ToUpper() }
                    
                    ServerLocatorService     = if (-not ($repHealth.Check -contains 'ServerLocatorService')) { "Unknown" }
                                               else { ($repHealth | Where-Object { $_.Check -eq 'ServerLocatorService' }).Result.ToUpper() }
                       
                    DagMembersUp     = if (-not ($repHealth.Check -contains 'DagMembersUp')) { "Unknown" }
                                       else { ($repHealth | Where-Object { $_.Check -eq 'DagMembersUp' }).Result.ToUpper() }

                    MonitoringService     = if (-not ($repHealth.Check -contains 'MonitoringService')) { "Unknown" }
                                            else { ($repHealth | Where-Object { $_.Check -eq 'MonitoringService' }).Result.ToUpper() }

                    ClusterNetwork     = if (-not ($repHealth.Check -contains 'ClusterNetwork')) { "Unknown" }
                                         else { ($repHealth | Where-Object { $_.Check -eq 'ClusterNetwork' }).Result.ToUpper() }

                    QuorumGroup     = if (-not ($repHealth.Check -contains 'QuorumGroup')) { "Unknown" }
                                      else { ($repHealth | Where-Object { $_.Check -eq 'QuorumGroup' }).Result.ToUpper() }
                    
                    FileShareQuorum     = if (-not ($repHealth.Check -contains 'FileShareQuorum')) { "Unknown" }
                                          else { ($repHealth | Where-Object { $_.Check -eq 'FileShareQuorum' }).Result.ToUpper() }

                    
                    <# The following are included only for info purposes. They do not get reported, so they will be commented out.
                    
                    DatabaseRedundancy     = if (-not ($repHealth.Check -contains 'DatabaseRedundancy')) { "Unknown" }
                                             else { ($repHealth | Where-Object { $_.Check -eq 'DatabaseRedundancy' }).Result.ToUpper() }

                    DatabaseAvailability     = if (-not ($repHealth.Check -contains 'DatabaseAvailability')) { "Unknown" }
                                               else { ($repHealth | Where-Object { $_.Check -eq 'DatabaseAvailability' }).Result.ToUpper() }

                    DBCopySuspended     = if (-not ($repHealth.Check -contains 'DBCopySuspended')) { "Unknown" }
                                          else { ($repHealth | Where-Object { $_.Check -eq 'DBCopySuspended' }).Result.ToUpper() }

                    DBCopyFailed     = if (-not ($repHealth.Check -contains 'DBCopyFailed')) { "Unknown" }
                                       else { ($repHealth | Where-Object { $_.Check -eq 'DBCopyFailed' }).Result.ToUpper() }

                    DBInitializing     = if (-not ($repHealth.Check -contains 'DBInitializing')) { "Unknown" }
                                         else { ($repHealth | Where-Object { $_.Check -eq 'DBInitializing' }).Result.ToUpper() }
                                        
                    DBDisconnected     = if (-not ($repHealth.Check -contains 'DBDisconnected')) { "Unknown" }
                                         else { ($repHealth | Where-Object { $_.Check -eq 'DBDisconnected' }).Result.ToUpper() }

                    DBLogCopyKeepingUp     = if (-not ($repHealth.Check -contains 'DBLogCopyKeepingUp')) { "Unknown" }
                                             else { ($repHealth | Where-Object { $_.Check -eq 'DBLogCopyKeepingUp' }).Result.ToUpper() }  

                    DBLogReplayKeepingUp     = if (-not ($repHealth.Check -contains 'DBLogReplayKeepingUp')) { "Unknown" }
                                               else { ($repHealth | Where-Object { $_.Check -eq 'DBLogReplayKeepingUp' }).Result.ToUpper() }

                    #>
                    
                    DAG = $dag.Name 
                }

                # Check for Failures/unknowns
                $notPassed = @($repHealth | Where-Object { $_.Result -ne 'Passed' }) | Select-Object Check,Result
                if ($notPassed.count -gt 0)
                {
                    foreach ($np in $notPassed)
                    {
                        if ($memberObj.$($np.Check) -ne $null)
                        {
                            $dagSummaryBag.Add("$dagMem in $($dag.Name) - '$($np.Check)' status is '$($np.Result)'")
                        }
                    }
                }

                # Add obj to report
                $dagMemberReportBag.Add($memberObj)
            }

            Write-Host "$($dag.Name.ToUpper()) - CHECKS COMPLETE" -f Green

        } # end of parallel foreach
    }

    Write-Host "`nDAG CHECKS COMPLETED" -f Green
    Write-Host "________________________`n`n"

    #####################
    # End of DAG Checks #
    #####################

    # Dispose the runspaces for the session
    $runSpace = Get-Runspace
    $runSpace | ForEach-Object { if ($_.RunspaceIsRemote -eq $true) { $_.Dispose() } }

    # Create and Sort the arrays
    $reportArray = @($serverReport | Sort-Object Server)
    $summaryArray = @($serverSummary | Sort-Object)
    $dagArray = @($dags | Sort-Object)
    $dagDbArray = @($dagDatabases | Sort-Object)
    $dagDbCopyArray = @($dagDBCopyReport | Sort-Object DatabaseName)
    $dagMemberArray = @($dagMemberReport | Sort-Object Server)

    ##############################
    # Begin HTML Report Buidling #
    ##############################
    
    # Set report generation time
    $reportTime = Get-Date

    # HTML Head
    $htmlhead = `
        "
        <html>
        <style>
        BODY {font-family: Ariall font-size: 8pt;}
        H1 {font-size: 16px;}
        H2 {font-size: 14px;}
        H3 {font-size: 12px;}
        TABLE {border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
        TH {border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
        TD {border: 1px solid black; padding: 5px;}
        td.pass {background: #7FFF00;}
        td.warn {background: #FFE600;}
        td.fail {background: #FF0000; color: #ffffff;}
        td.info {background: #85D4FF;}
        </style>
        <body>
        <h1 align=""center"">Exchange Server Health Check Report</h1>
        <h3 align=""center"">Generated: $reportTime</h3>
        "

    # Server Summary HTML
    if ($serverSummary.Count -eq 0)
    {
        $serverSummaryHTML = `
            "
            <h3>Exchange Server Health Check Summary</h3>
            <p>No Exchange Server Health errors or warnings.</p>
            "
        $serverSummary.Add("No Exchange Server Health errors or warnings.")
    }
    else
    {
        $serverSummaryHTML = `
            "
            <h3>Exchange Server Health Check Summary</h3>
            <p>The following server errors and warnings were detected.</p>
            <p>
            <ul>
            "
        foreach ($srv in $serverSummary)
        {
            $serverSummaryHTML += "<li>$srv</li>"
        }
        
        $serverSummaryHTML += "</ul></p>"
    }
    
    # DAG Summary HTML
    if ($dagSummary.count -eq 0)
    {
        $dagSummaryHTML = `
            "
            <h3>Database Availability Group Health Check Summary</h3>
            <p>No Exchange DAG Errors or Warnings.</p>
            "
        $dagSummary.Add("No Exchange DAG Errors or Warnings.")
    }
    else
    {
        $dagSummaryHTML = `
            "
            <h3> Database Availability Group Health Check Summary</h3>
            <p>The following DAG Errors and Warnings were detected.</p>
            <p>
            <ul>
            "
        foreach ($srv in $dagSummary)
        {
            $dagSummaryHTML += "<li>$srv</li>"
        }

        $dagSummaryHTML += "</ul></p>"
    }

    # Server Health Check Report Table
    $serverHealthHTMLtable = "" # Clearing

    # Table Header
    $htmlTableHeader = `
        "
        <h3>Exchange Server Health</h3>
        <p>
        <table>
        <tr>
        <th>Server</th>
        <th>Site</th>
        <th>DNS</th>
        <th>Ping</th>
        <th>Uptime</th>
        <th>Client Access Server Role Services</th>
        <th>Hub Transport Server Role Services</th>
        <th>Mailbox Server Role Services</th>
        <th>Transport Queue</th>
        <th>MB DBs Mounted</th>
        <th>Mail Flow Test</th>
        <th>MAPI Test</th>
        </tr>
        "
    $serverHealthHTMLtable = $serverHealthHTMLtable + $htmlTableHeader

    # Begin Adding each server to the HTML Table
    foreach ($svr in $reportArray)
    {
        $htmlTableRow  = "<tr>"
        $htmlTableRow += "<td>$($svr.Server)</td>"
        $htmlTableRow += "<td>$($svr.Site)</td>"
        $htmlTableRow += (New-ServerHTMLTableCell -result $svr.DNS)
        $htmlTableRow += (New-ServerHTMLTableCell -result $svr.Ping)
        $htmlTableRow += (New-ServerHTMLTableCell -result $svr.'Uptime (hrs)' -isINT)
        $htmlTableRow += (New-ServerHTMLTableCell -result $svr.'CASR Services')
        $htmlTableRow += (New-ServerHTMLTableCell -result $svr.'HTSR Services')
        $htmlTableRow += (New-ServerHTMLTableCell -result $svr.'MBSR Services')
        $htmlTableRow += (New-ServerHTMLTableCell -result $svr.'Transport Queue')
        $htmlTableRow += (New-ServerHTMLTableCell -result $svr.'MB DBs Mounted')
        $htmlTableRow += (New-ServerHTMLTableCell -result $svr.'Mail Flow Test')
        $htmlTableRow += (New-ServerHTMLTableCell -result $svr.'MAPI Test')
        $htmlTableRow += "</tr>"

        # Add the server to the table
        $serverHealthHTMLtable = $serverHealthHTMLtable + $htmlTableRow
    }

    # Finish the Server health table
    $serverHealthHTMLtable = $serverHealthHTMLtable + "</table></p>"

    # DAG Health HTML Tables
    $dagHealthHTMLtable = ""

    # Go through each DAG
    foreach ($dag in $dagArray)
    {
        # Reset html variables
        $dagSumHTML = ""
        $dagDetHTML = ""
        $dagMemHTML = ""

        # Strings for use in their respective HTML Report
        $dagSummaryIntro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Health Summary:</p>"
        $dagDetailsIntro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Health Details:</p>"
        $dagMembersIntro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Member Health:</p>"

        # Gather DAG Specific details
        $dagDBs      = $dagDbArray | Where-Object { $_.DAG -eq $dag.Name }
        $dagDBCopies = $dagDbCopyArray | Where-Object { $_.DAG -eq $dag.Name }
        $dagMembers  = $dagMemberArray | Where-Object { $_.DAG -eq $dag.Name }

        # DAG HTML Health Summary Table Header
        $dagSumHTMLheader = `
            "
            <p>
            <table>
            <tr>
            <th>Database</th>
            <th>Mounted On</th>
            <th>Total Copies</th>
            <th>Healthy Copies</th>
            <th>Unhealthy Copies</th>
            <th>Healthy Queues</th>
            <th>Unhealthy Queues</th>
            </tr>
            "
        $dagSumHTML = $dagSummaryIntro + $dagSumHTMLheader

        # Go through each DB and add to the table
        foreach ($db in $dagDbs)
        {
            $htmlTableRow  = "<tr>"
            $htmlTableRow += "<td>$($db.Database)</td>"
            $htmlTableRow += "<td>$($db.'Mounted On')</td>"
            $htmlTableRow += "<td>$($db.'Total Copies')</td>"
            $htmlTableRow += (New-DagHTMLTableCell -result @($db.'Healthy Copies',$db.'Total Copies') -Summary -healthy)
            $htmlTableRow += (New-DagHTMLTableCell -result @($db.'Unhealthy Copies',$db.'Total Copies') -Summary)
            $htmlTableRow += (New-DagHTMLTableCell -result @($db.'Healthy Queues',$db.'Total Copies') -Summary -healthy)
            $htmlTableRow += (New-DagHTMLTableCell -result @($db.'Unhealthy Queues',$db.'Total Copies') -Summary)
            $htmlTableRow += "</tr>"

            $dagSumHTML = $dagSumHTML + $htmlTableRow
        }

        # Finish DAG Health Summary Table
        $dagSumHTML = $dagSumHTML + "</table></p>"
        $dagHealthHTMLtable = $dagHealthHTMLtable + $dagSumHTML

        # DAG Details HTML Header
        $dagDetHTMLheader = `
            "
            <p>
            <table>
            <tr>
            <th>Database Copy</th>
            <th>Database Name</th>
            <th>Mailbox Server</th>
            <th>Status</th>
            <th>Copy Queue</th>
            <th>Replay Queue</th>
            </tr>
            "
        $dagDetHTML = $dagDetailsIntro + $dagDetHTMLheader

        # Go through each DB Copy
        foreach ($dbCopy in $dagDBCopies)
        {
            $htmlTableRow  = "<tr>"
            $htmlTableRow += "<td>$($dbCopy.'Database Copy')</td>"
            $htmlTableRow += "<td>$($dbCopy.'Database Name')</td>"
            $htmlTableRow += "<td>$($dbCopy.'Mailbox Server')</td>"
            $htmlTableRow += (New-DagHTMLTableCell -result $dbCopy.Status -Details)
            $htmlTableRow += (New-DagHTMLTableCell -result $dbCopy.'Copy Queue' -Details -isINT)
            $htmlTableRow += (New-DagHTMLTableCell -result $dbCopy.'Replay Queue' -Details -isINT)
            $htmlTableRow += "</tr>"

            $dagDetHTML = $dagDetHTML + $htmlTableRow
        }

        # Finish DAG Details Table
        $dagDetHTML = $dagDetHTML +  "</table></p>"
        $dagHealthHTMLtable = $dagHealthHTMLtable + $dagDetHTML

        # DAG Members HTML Header
        $dagMemHTMLheader = `
            "
            <p>
            <table>
            <tr>
            <th>Server</th>
            <th>Cluster Service</th>
            <th>Replay Service</th>
            <th>Active Manager</th>
            <th>Tasks RPC Listener</th>
            <th>TCP Listener</th>
            <th>Server Locator Service</th>
            <th>DAG Members Up</th>
            <th>Cluster Network</th>
            <th>Quorum Group</th>
            <th>FileShare Quorum</th>
            </tr>
            "
        $dagMemHTML = $dagMembersIntro + $dagMemHTMLheader

        # Go through each DAG Member
        foreach ($dagMem in $dagMembers)
        {
            $htmlTableRow  = "<tr>"
            $htmlTableRow += "<td>$($dagMem.Server)</td>"
            $htmlTableRow += (New-DagHTMLTableCell -result $dagMem.ClusterService -Members)
            $htmlTableRow += (New-DagHTMLTableCell -result $dagMem.ReplayService -Members)
            $htmlTableRow += (New-DagHTMLTableCell -result $dagMem.ActiveManager -Members)
            $htmlTableRow += (New-DagHTMLTableCell -result $dagMem.TasksRpcListener -Members)
            $htmlTableRow += (New-DagHTMLTableCell -result $dagMem.TcpListener -Members)
            $htmlTableRow += (New-DagHTMLTableCell -result $dagMem.ServerLocatorService -Members)
            $htmlTableRow += (New-DagHTMLTableCell -result $dagMem.DagMembersUp -Members)
            $htmlTableRow += (New-DagHTMLTableCell -result $dagMem.ClusterNetwork -Members)
            $htmlTableRow += (New-DagHTMLTableCell -result $dagMem.QuorumGroup -Members)
            $htmlTableRow += (New-DagHTMLTableCell -result $dagMem.FileShareQuorum -Members)
            $htmlTableRow += "</tr>"

            $dagMemHTML = $dagMemHTML + $htmlTableRow
        }

        # Finish DAG Members Table
        $dagMemHTML = $dagMemHTML + "</table></p>"
        $dagHealthHTMLtable = $dagHealthHTMLtable + $dagMemHTML
    }
    # End DAG HTML Tables

    # Finish building the HTML Report
    $htmlTail = "</body></html>"
    $htmlReport = $htmlHead + $serverSummaryHTML + $dagSummaryHTML + $serverHealthHTMLtable + $dagHealthHTMLtable + $htmlTail

    #########################
    # END HTML TABLE REPORT #
    #########################

    ##################
    # Console Output #
    ##################

    # Exchange Summary
    Write-Host "`n====================================="
    Write-Host "Exchange Server Health Check Summary"
    Write-Host "=====================================`n"
    $serverSummary | Sort

    # Create space
    Write-Host "`n`n"

    # DAG Summary
    Write-Host "`n================================================="
    Write-Host "Database Availability Group Health Check Summary"
    Write-Host "=================================================`n"
    $dagSummary | Sort
    
    # create Space
    Write-Host "`n`n"

    #########################
    # End of Console Output #
    #########################

    $SendEmail = $true
    if ($SendEmail)
    {
        Send-MailMessage @smtpsettings -Body $htmlReport -BodyAsHtml
    }
}
