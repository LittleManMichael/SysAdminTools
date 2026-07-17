#########################################################
#                                                       #
# These Functions are for SysAdmin Purposes             #
#                                                       #
#########################################################
Set-PowerCLIConfiguration -ParticipateInCeip $false -Scope AllUsers  -Confirm:$false  |out-null
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Scope AllUsers -Confirm:$false |out-null
Set-PowerCLIConfiguration -DisplayDeprecationWarnings $false -Scope AllUsers -Confirm:$false |out-null
Set-PowerCLIConfiguration -DefaultVIServerMode Multiple -confirm:$false | Out-Null

function Get-RemoteCU {
    param (
        [Parameter(Mandatory=$true)]
        [string]$PC
    )
    if (-not (test-connection -ComputerName $PC -count 3 -Quiet -BufferSize 16)) { 
        throw "Machine is offline: $PC"
    }

    try {
        Get-HotFix -ComputerName $PC | Where-Object { $_.Description -like "*Cumulative Update*" } | 
        Select-Object HotFixID, Description, InstalledOn | Sort-Object InstalledOn -Descending
    }
    catch { 
        Write-Warning "Failed to gather info from $PC...`n $($_.Exception)"
    }
}


Function Disable-StaleAdmins {
<#
 .Synopsis
  Disable ALL Stale Admin accounts that have been found via the Get-StaleAdmins script. 

 .Description
  Administrator accounts will be disabled if found stale. This script will disable them.

 .Example
    Disable-StaleAdmins
    # Runs the Script, Disabling all Stale Admin Accounts.
#>


#--------------------------------------------------------------------------------------------------------
# Created 08 Oct 2022 by Michael Sprous
#
# Updated: 16 Oct 2025 by Michael Sprous
# - Overhaul of script to allow functioning for -AD accounts.
# - Updated the reporting output. 
#
# Updated: 4 Jan 2026 by Michael Sprous
# - Updated output and reporting methods.
#
# Updated: 29 May 2026 by Michael Sprous
# - Rewrite of the script entirely.
#--------------------------------------------------------------------------------------------------------

[cmdletbinding()]
param( [Parameter(Mandatory=$false)] [string]$IncidentNumber )

# Verify Running as Admin
if ($env:USERNAME -notlike '*-admin') {
    throw "Insufficient permissions! Please run this script from your -Admin account.."
}

# Verify -AD Credentials
Write-Host "Beginning -AD credential check..." -ForegroundColor Yellow
$credTested = $false
$adCred = $null
$tstcnt = 0
while ($credTested -eq $false) {
    if ($tstcnt -eq 5) { throw "Too many attempts... Please try again.." }

    $cred = Get-Credential -Message "Enter your -AD account credentials."
    try {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        $testCred = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain)

        if ($testCred.ValidateCredentials($cred.GetNetworkCredential().UserName, $cred.GetNetworkCredential().Password)) {
            Write-Host "-AD credentials verified." -ForegroundColor Green
            $adCred = $cred
            $credtested = $true 
        }
        else {
            Write-Warning "-AD credentials failed. Please try again."
            $tstcnt++
        }
    }
    catch {
        Write-Warning "An error occurred during credential validation.. It may be a network issue or your account is locked out. 
        $_"
        $tstcnt++

    }
}

# Request and verify incident number
if (-not $IncidentNumber) {
    Write-Host "Requesting Incident Number.." -ForegroundColor Yellow
    $IncidentNumber = Read-Host "Please enter a valid ServiceNow incident number."
}
while ($IncidentNumber -notmatch '^[0-9]{4}|^[0-9]{5}|^[0-9]{6}|^[0-9]{7}|^[0-9]{8}|^[0-9]{9}$') {
    $IncidentNumber = Read-Host "Please enter a valid ServiceNow incident number."
}
$IncidentNumber = "INC" + $IncidentNumber

# Setup Initial Variables
$LogPath = "\\newton\admin\scripts\Monthly Scripts\Stale Admin Accounts\Logs"
$Results = [System.Collections.Generic.List[Object]]::new()
$CurrentDateStr = Get-Date -Format 'dd MMM yyyy'
$NewDescription = "Disabled due to inactivity on $CurrentDateStr - $IncidentNumber"
$ExpiredDescription = $NewDescription -replace 'Disabled','Expired'
$Days = 90
$ccnt = 0

# Gather User List Path
$LatestLogDir = Get-ChildItem -Path $LogPath -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $LatestLogDir) { throw "No log directories found in '$LogPath'." }
$TargetAdminFile = Join-Path $LatestLogDir.FullName "Admins.txt"
if (-not (Test-Path $TargetAdminFile)) { throw "Admins file not found at '$TargetAdminFile'." }
$Admins = Get-Content $TargetAdminFile | Where-Object { $_ -match '\S' -and $_ -notlike '*SVC-*' -and $_ -notlike '*-SVC*' }
if (-not $Admins) { Write-Host "Admins file '$TargetAdminFile' is empty. No accounts to process!" -f Green ; return }
$tcnt = $Admins.Count 

# Gather server information
$TargetDC        = (Get-ADDomainController -Discover -ForceDiscover).HostName
#$TargetMBX       = (Get-ADComputer -Filter "Name -like 'N020MBX*'" | Select-Object -First 1).DNSHostName
#$TargetProfileFP = (Get-ADComputer -Filter "Name -like 'N020FP0*'" | Where-Object { $_.Name -match '3' -or $_.Name -match '7' } | Select-Object -First 1).DNSHostName
#$TargetHomeFP    = (Get-ADComputer -Filter "Name -like 'N020FP0*'" | Where-Object { $_.Name -match '4' -or $_.Name -match '8' } | Select-Object -First 1).DNSHostName

#- Output the settings used
Write-Host "The following settings will be used: 
-- SERVERS --
Active Directory : $TargetDC
Profile Folder   : Not Used
Home Folder      : Not Used
Exchange Actions : Not Used

-- INFO --
Incident Number  : $IncidentNumber
Userlist Path    : $TargetAdminFile
Total Users      : $tcnt

-- EXEMPTIONS --
None Found.

NOTE: Please cancel if anything looks wrong.
"

Start-Sleep -Seconds 3
Write-Host "`nBeginning Account Disablement..."

# Connect to the exchange server
$RemoteSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$($TargetMBX)/powershell
Import-PSSession $RemoteSession

# Begin Account processing loop
foreach ($SamName in $Admins) {

    # Adjust the progressbar
    [int]$pct = ($ccnt / $tcnt) * 100
    Write-Progress -Activity "Going through Stale Admins..." -Status "$ccnt/$tcnt - $SamName" -PercentComplete $pct
    $ccnt++

    # Gather and verify account exists in ADUC
    $Admin = Get-ADUser -Identity $SamName -Properties * -ErrorAction SilentlyContinue
    if (-not $Admin) {
        Write-Warning "Account '$SamName' not found in ADUC. Skipping."
        $Results.Add([PSCustomObject]@{ SamAccountName = $SamName ; Status = "Not Found in ADUC" })
        continue
    }

    # Date Calculation
    $TargetDate = if ($Admin.LastLogonDate) { $Admin.LastLogonDate } else { $Admin.whenCreated }
    $DaysInactive = ((Get-Date) - $TargetDate).Days 

    # Gather AD Security Groups
    $FriendlyGroups = if ($Admin.MemberOf) { ($Admin.MemberOf | ForEach-Object { ($_ -split ',')[0] -replace '^CN=',''}) -join '; ' } else { "None" }

    # Create LogEntry object
    $LogEntry = [PSCustomObject]@{
        SamAccountName = $Admin.SamAccountName
        DaysInactive = $DaysInactive
        ActionTaken = "None."
        Status = "Skipped"
        OriginalGroups = $FriendlyGroups
    }

    # EXCLUSIONS
    # Currently no exclusions found

    # Complete DC Actions
    try {
        Invoke-Command -ComputerName $TargetDC -Credential $adCred -ScriptBlock {
            param($AdminIdentity, $NewAdminDescription, $AdminMemberOf)

            # Disable Account
            Disable-ADAccount -Identity $AdminIdentity
            
            # Change Description
            Set-ADUser -Identity $AdminIdentity -Description $NewAdminDescription

            # Remove all groups (Domain Users is natively excluded
            if ($AdminMemberOf) { 
                $AdminMemberOf | ForEach-Object {
                    Remove-ADGroupMember -Identity $_ -Members $AdminIdentity -Confirm:$false -ErrorAction SilentlyContinue
                }
            }

            # Move to Disabled OU
            $ADAdmin = Get-ADUser -Identity $AdminIdentity -Properties DistinguishedName
            $CurrentOU = ([ADSI]"LDAP://$($ADAdmin.DistinguishedName)").Parent -replace '^LDAP://',''
            $DisabledOU = Get-ADOrganizationalUnit -Filter "Name -like '*disabled*'" -SearchBase $CurrentOU -SearchScope OneLevel | Select-Object -First 1
            if ($DisabledOU) {
                Move-ADObject -Identity $ADAdmin.DistinguishedName -TargetPath $DisabledOU.DistinguishedName -ErrorAction SilentlyContinue
            }
        } -ArgumentList $Admin.SamAccountName, $NewDescription, $Admin.MemberOf

        # Update logentry object
        $LogEntry.Status = 'DISABLED'
        $LogEntry.ActionTaken = 'DISABLED'
    }
    catch {
        Write-Warning "Failed to Disable '$SamName' (They will be skipped): $($_.Exception.Message)"
        $LogEntry.Status = "Failed: $($_.Exception.Message)"
        continue
    }

    # Add results to the results table
    $Results.Add($LogEntry)
}

# Finish the progressbar
Write-Progress -Activity "Going through Stale Admins..." -Completed

# Final reporting
$OutCsv = Join-Path (Split-Path $TargetAdminFile) "Disable-StaleAdmins-Report-$(Get-Date -f 'yyyyMMddHHmmss').csv"
$Results | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8 -Force

$Results | Format-Table -AutoSize
Write-Host "`nProcess Complete. Report saved to: $OutCsv" -ForegroundColor Green
Write-Host "`n=== IMPORTANT ===" -ForegroundColor Red -BackgroundColor Black
Write-Host "You will need to push the report to Gold via iSAFE. 
Reports are to be emailed together (N/S/G) to PMO CyberSecurity.
Reports are expected every Monday by PMO CyberSecurity.
"

# Logging of script
$LogPath = "\\newton\admin\SysAdmin Powershell Module\Scripts\BattleRythm\Logs\Log_BattleRhythm.csv"
Add-Content -Path $LogPath -Value "`Disable-StaleAdmins,$(Get-Date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami)"

}

Function Disable-StaleUsers {

<#
 .Synopsis
  Disable ALL Stale Admin accounts that have been found via the Get-StaleAdmins script. 

 .Description
  User accounts will be disabled if found stale. This script will disable them.

 .Example
    Disable-Staleusers
    # Runs the Script, Disabling all Stale user Accounts.
#>


#--------------------------------------------------------------------------------------------------------
# Created 08 Oct 2022 by Michael Sprous
#
# Updated: 16 Oct 2025 by Michael Sprous
# - Overhaul of script to allow functioning for -AD accounts.
# - Updated the reporting output. 
#
# Updated: 10 Mar 2026 by Michael Sprous
# - Added RT_Ops exemption
#
# Update: 28 Mar 2026 by Michael Sprous
# - Updated the LLTS to show as when the account was created if the user never actually logged on. This prevents excluded accounts from being missed if they never logged on. 
#
# Updatee 4 June 2026 by Michael Sprous
# - Rewrite of the script to improve output and reporting.
#
#
#--------------------------------------------------------------------------------------------------------


[cmdletbinding()]
param(
    [Parameter(Mandatory=$false)] [string]$IncidentNumber 
)


#- Verify Running as -Admin 
if ($env:USERNAME -notlike '*-admin') {
    throw "Insufficient Permissions! Please run this script from your -Admin account.."
}

#- Verify -AD Credentials
Write-Host "Beginning -AD credential check..." -ForegroundColor Yellow
$credTested = $false
$adCred = $null 
$tstcnt = 0
while ($credTested -eq $false) {
    if ($tstcnt -eq 5) { throw "Too many attempts... Pleast run script again" }

    $cred = Get-Credential -Message "Enter our -AD account credentials."
    try {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        $testCred = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain)

        if ($testCred.ValidateCredentials($cred.GetNetworkCredential().UserName, $cred.GetNetworkCredential().Password)) {
            Write-Host "-AD Credentials Verified" -ForegroundColor Green
            $adCred = $cred
            $credtested = $true
        }
        else { 
            Write-Warning "-AD Credentials failed. Please try again." 
            $tstcnt++
        }
    }
    catch { 
        Write-Warning "An Error occurred during credential validation. It may be a network issue or locked account. $_" 
        $tstcnt++
    }
}

#- Request and Verify Incident Number
if (-not $IncidentNumber) {
    Write-Host "Requesting Incident Number..." -ForegroundColor Yellow
    $IncidentNumber = Read-Host "Please Enter a Valid ServiceNow Incident Number"
}
While ($IncidentNumber -notmatch '^[0-9]{4}$|^[0-9]{5}$|^[0-9]{6}$|^[0-9]{7}$|^[0-9]{8}$|^[0-9]{9}$') {
        $IncidentNumber = Read-Host "Please Enter a Valid ServiceNow Incident Number"
}
$IncidentNumber = "INC" + $IncidentNumber

#- Setup initial variables
$LogPath = "\\newton\admin\scripts\monthly scripts\stale user accounts\logs"
$Days = 90
$Results = [System.Collections.Generic.List[object]]::new()
$CurrentDateStr = Get-Date -Format 'dd MMM yyyy'
$NewDescription = "Disabled due to inactivity on $CurrentDateStr - $IncidentNumber"
$ExpiredDescription = $NewDescription -replace "Disabled","Expired"

#- Gather User list path
$LatestLogDir = Get-ChildItem -Path $LogPath -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $LatestLogDir) { throw "No log directories found in '$LogPath'." }
$TargetUserFile = Join-Path $LatestLogDir.FullName "Users.txt"
if (-not (Test-Path $TargetUserFile)) { throw "User file not found at '$TargetUserFile'." }
$Users = Get-Content $TargetUserFile | Where-Object { $_ -match '\S' }
if (-not $Users) { Write-Host "User file '$TargetUserFile' is empty. No Accounts to process!" -f Green ; return }
$tcnt = $Users.Count 
$ccnt = 0

#- Gather Server informaiton
$TargetDC = (Get-ADDomainController -Discover -ForceDiscover).HostName
$TargetMBX = (Get-ADComputer -Filter "Name -like 'N020MBX*'" | Select -First 1).DNSHostName
$TargetProfileFP = (Get-ADComputer -Filter "Name -like 'N020FP0*'" | Where { $_.Name -match '3' -or $_.Name -match '7' }| Select -First 1).DNSHostName
$TargetHomeFP = (Get-ADComputer -Filter "Name -like 'N020FP0*'" | Where { $_.Name -match '4' -or $_.Name -match '8' }| Select -First 1).DNSHostName

#- Output the settings used
Write-Host "The following settings will be used: 
-- SERVERS --
Acitve Directory    : $TargetDC
Profile Folder      : $TargetProfileFP
Home Folder         : $TargetHomeFP
Exchange Actions    : $TargetMBX

-- INFO --
Incident Number     : $IncidentNumber
Userlist Path       : $TargetUserFile
Total Users         : $tcnt

-- EXEMPTIONS --
RT_OPS              : Exempt for 1 year of inactivity.
365 Exemption Group : Exempt for 1 year of inactivity.

NOTE: Please cancel if anything looks wrong.
"

Start-sleep -Seconds 3
Write-Host "`n Beginning Account Disablement..."

#- Connect to the exchange session
$RemoteSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$($TargetMBX)/powershell
Import-PSSession $RemoteSession

#- BEGIN ACCOUNT PROCESSING LOOP
foreach ($SamName in $Users) {
    
    #- Adjust the progress bar
    [int]$pct = ($ccnt / $tcnt) * 100
    Write-Progress -Activity "Going through Stale Users..." -Status "$ccnt/$tcnt - $SamName" -PercentComplete $pct
    $ccnt++
    
    #- Gather and Verify user exists in ADUC
    $User = Get-ADUser -Identity $SamName -Properties * -ErrorAction SilentlyContinue
    if (-not $User) {
        Write-Warning "User '$SamName' not found in ADUC. Skipping."
        $Results.Add([PSCustomObject]@{ SamAccountName = $SamName ; Status = "Not Found in ADUC" })
        continue
    }

    #- Date Calculation
    $TargetDate = if ($User.LastLogonDate) { $User.LastLogonDate } else { $User.WhenCreated }
    $DaysInactive = ((Get-Date) - $TargetDate).Days
    
    $FriendlyGroups = if ($User.MemberOf) { ($User.MemberOf | ForEach-Object { ($_ -split ',')[0] -replace '^CN=',''}) -join '; ' } else { "None" } 
    
    $LogEntry = [PSCustomObject]@{
        SamAccountName = $User.SamAccountName 
        DaysInactive   = $DaysInactive
        ActionTaken    = "None."
        ProfileFolder  = "N/A"
        HomeFolder     = "N/A"
        ExchMailbox    = "N/A"
        Status         = "Skipped"
        OriginalGroups = $FriendlyGroups
    }

    # --- EXCLUSIONS --- 
    #- Check RT_OPS (Excempt for under 1 year / 379 days)
    $RTOpsGroup = (Get-ADGroup "RT_OPS" -ErrorAction SilentlyContinue).DistinguishedName
    $365Exmpt = (Get-ADGroup "365 Exemption Group" -ErrorAction SilentlyContinue).DistinguishedName
    if ( ($RTOpsGroup -and ($User.MemberOf -contains $RTOpsGroup) -and ($DaysInactive -lt 379)) -or ($365Exmpt -and ($User.MemberOf -contains $365Exmpt) -and ($DaysInactive -lt 379)) ) {
        Write-Host "Skipping $SamName (RT_OPS or 365 Exemption Group Exempt, inactive $DaysInactive days)" -ForegroundColor DarkGray
        $LogEntry.ActionTaken = "Exempt - RT_OPS or 365 Exemption Group"
        $Results.Add($LogEntry)
        continue
    }

    # --- 180+ Days Logic (194 buffer) ---
    if ($DaysInactive -ge 194) {
        Write-Host "Disabling $SamName ($DaysInactive days)..." -ForegroundColor Yellow

        # Begin trying to disable
        try {
            #- 1. Execute all AD changes remotely on the DC
            Invoke-Command -ComputerName $TargetDC -Credential $adCred -ScriptBlock {
                param($UserIdentity, $NewUserDesc, $UserMemberOf)

                #- 1a. Disable account and update description
                Disable-ADAccount -Identity $UserIdentity 
                Set-ADUser -Identity $UserIdentity -Description $NewUserDesc

                #- 1b. Remove from all groups (Domain Users is natively exlcuded
                if ($UserMemberOf) {
                    $UserMemberOf | ForEach-Object {
                        Remove-ADGroupMember -Identity $_ -Members $UserIdentity -Confirm:$false -ErrorAction SilentlyContinue
                    }
                }

                #- 1c. Move to local 'Disabled' OU
                $ADUser = Get-ADUser -Identity $UserIdentity -Properties DistinguishedName
                $CurrentOU = ([ADSI]"LDAP://$($User.DistinguishedName)").Parent -replace '^LDAP://', ''
                $DisabledOU = Get-ADOrganizationalUnit -Filter "Name -like '*disabled*'" -SearchBase $CurrentOU -SearchScope OneLevel | Select-Object -First 1
                if ($DisabledOU) {
                    Move-ADObject -Identity $ADUser.DistinguishedName -TargetPath $DisabledOU.DistinguishedName -ErrorAction SilentlyContinue
                }
            } -ArgumentList $User.SamAccountName, $NewDescription, $User.MemberOf 

            $LogEntry.ActionTaken = "Disabled"
        }
        catch {
            Write-Warning "Failed to Disable '$SamName'(They will be skipped): $($_.Exception.Message)"
            $LogEntry.Status = "Failed: $($_.Exception.Message)"
            continue
        }
        
        # Begin Renaming Profile Folder
        try {    
            #- 1. Gather Profile Path    
            $pPath = ($User.ProfilePath -split '\\' | Select-Object -Last 3) -join '\'

            #- 2. Execute all Profile Server changes
            Invoke-Command -ComputerName $TargetProfileFP -ScriptBlock {
                param($ProfilePath, $Incident) 

                #- 2a. Find the profile path
                $ProfilePath = Join-Path "E:" "$ProfilePath.v6"
                $testPath = Test-Path -Path $ProfilePath
                $newName = "$($ProfilePath | Split-Path -Leaf)-Disabled-$($Incident)"
                
                #- 2b. Rename the Path
                if ($testPath) {
                    try {
                        Rename-Item -Path $ProfilePath -NewName $newName
                    }
                    catch {
                        takeown.exe /f $ProfilePath /a /r /d y > $null
                        icacls.exe $ProfilePath /grant $env:USERNAME /t /c /q > $null
                        Rename-Item -Path $ProfilePath -NewName $newName
                    }
                }
            } -ArgumentList $pPath, $IncidentNumber

            $LogEntry.ProfileFolder = "Renamed"
        }
        catch {
            Write-Warning "Failed to rename the Profile Folder for '$SamName': $($_.Exception.Message)"
            $LogEntry.ActionTaken   = "Disabled"
            $LogEntry.ProfileFolder = "Failed: $($_.Exception.Message)"
            $LogEntry.Status        = "Success with Errors"
        }

        # Begin Renaming the Home Folder
        try {    
            #- 1. Gather Home Path    
            $hPath = ($User.ProfilePath -split '\\' | Select-Object -Last 3) -join '\'

            #- 2. Execute all Profile Server changes
            Invoke-Command -ComputerName $TargetHomeFP -ScriptBlock {
                param($HomePath, $Incident) 

                #- 2a. Find the profile path
                $HomePath = Join-Path "E:" "$Home"
                $testPath = Test-Path -Path $HomePath
                $newName = "$($HomePath | Split-Path -Leaf)-Disabled-$($Incident)"
                
                #- 2b. Rename the Path
                if ($testPath) {
                    try {
                        Rename-Item -Path $HomePath -NewName $newName
                    }
                    catch {
                        takeown.exe /f $HomePath /a /r /d y > $null
                        icacls.exe $HomePath /grant $env:USERNAME /t /c /q > $null
                        Rename-Item -Path $HomePath -NewName $newName
                    }
                }
            } -ArgumentList $hPath, $IncidentNumber

            $LogEntry.HomeFolder = "Renamed"
        }
        catch {
            Write-Warning "Failed to rename the Home Folder for '$SamName': $($_.Exception.Message)"
            $LogEntry.ActionTaken = "Disabled"
            $LogEntry.HomeFolder  = "Failed: $($_.Exception.Message)"
            $LogEntry.Status      = "Success with Errors"
        }

        # Begin Disabling the Mailbox
        try {
            #- 1. Disable Mailbox
            Set-Mailbox -Identity $User.SamAccountName -EmailAddressPolicyEnabled $False -HiddenFromAddressListsEnabled $True
            Set-CASMailbox -Identity $User.SamAccountName -ActiveSyncEnabled $False -ImapEnabled $False -MAPIEnabled $False -OWAEnabled $False -PopEnabled $False

            $LogEntry.ExchMailBox = "Disabled"
        }
        catch {
            Write-Warning "Failed to disable the Exchange Mailbox for '$SamName': $($_.Exception.Message)"
            $LogEntry.ActionTaken = "Disabled"
            $LogEntry.ExchMailBox = "Failed: $($_.Exception.Message)"
            $LogEntry.Status      = "Success with Errors"
        }

        if ($LogEntry.Status -ne "Success with Errors") { $LogEntry.Status = "Success" }
    }

    #--- 90+ Days Logic (104 buffer) ---
    elseif ($DaysInactive -ge 104) {
        Write-Host "Expiring $SamName ($DaysInactive days)..." -ForegroundColor Yellow

        #- Begin Expiration
        try {
            #- 1. Attempt to Expire account in ADUC
            Invoke-Command -ComputerName $TargetDC -Credential $adCred -ScriptBlock {
                param($UserIdentity,$NewUserDesc)
                
                #- 1a. Expire the account and update the description
                Set-ADAccountExpiration -Identity $UserIdentity -DateTime (Get-Date).Date 
                Set-ADUser -Identity $UserIdentity -Description $NewUserDesc
            } -ArgumentList $User.SamAccountName, $ExpiredDescription

            $LogEntry.ActionTaken = "Expired"
            $LogEntry.Status
        }
        catch {
            $LogEntry.Status = "Failed $($_.Exception.Message)"
            Write-Warning "Failed to expire '$SamName': $($_.Exception.Message)"
        }

    }

    #- Finally, add the results to the Results Table
    $Results.Add($LogEntry)
}

# Finish the progress bar
Write-Progress -Activity "Going Through Stale Users..." -Completed

#- FINAL REPORTING
$OutCsv = Join-Path (Split-Path $TargetUserFile) "Disable-StaleUsers-Report-$(Get-Date -f 'yyyyMMddHHmmss').csv" 
$Results | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8 -Force

$Results | Format-Table -AutoSize
Write-Host "`nProcess Complete. Report saved to: $OutCsv" -ForegroundColor Green
Write-Host "`n=== IMPORTANT ===" -f Red -b Black
Write-Host "You will need to push the report to Gold via iSAFE.
Reports are to be emailed together (N/S/G) to PMO CyberSecurity
Reports are expected every Monday by PMO CyberSecurity
"

# Logging of script
$logpath = "\\newton\admin\SysAdmin Powershell Module\Scripts\BattleRythm\Logs\Log_BattleRhythm.csv"
Add-Content -Path $logpath -Value "`Disable-StaleUsers,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami)"
}

function loop-checks {
[int]$oddeven = 1
start-process powershell.exe -argumentlist "-noexit check-smtpQ"
    Do {
        [int]$Xmin = 120
        $nextrun = [datetime]$currentTime = (get-date).AddMinutes($Xmin)
        start-process powershell.exe -argumentlist "-noexit whats-up-doc"
        if ($oddeven % 2 -ne 0) {
            start-process powershell.exe -argumentlist "-noexit test-exchangeserverhealth -sendemail"
            }
        $oddeven = $oddeven + 1
        do {
        $currentTime = (get-date)
        $countdown = New-TimeSpan -start (get-date) -end $nextrun

        cls
        Write-progress -Activity "Time till next checks:" -status "Countdown - $("{0:D2}" -f ($countdown.hours))`:$("{0:D2}" -f ($countdown.minutes))`:$("{0:D2}" -f ($countdown.seconds))"
        start-sleep 1

        } Until ($($countdown.totalminutes) -le 0)

    Write-progress -Activity "Time till next checks:" -completed
    } until (1 -eq 2)
}

Function Clear-XDCLicenses {
$xdcservers = @(get-adcomputer -filter "name -like '*0xdc*'" |select name).name
$path = '$\XenDesktop\Licensing\LS\resource\cache'

Foreach ($xdc in $xdcservers) {
Write-host "Creating CIM session to $xdc" -ForegroundColor Green
$cim = New-CimSession -ComputerName $xdc

Write-host "Finding the $xdc XenDesktop Drive Letter" -ForegroundColor Green
$Xendrive = get-volume -CimSession $cim -FileSystemLabel "XenDesktop"
$XendriveLetter = $Xendrive.DriveLetter

Write-host "Stopping the $xdc Citrix Licensing service" -ForegroundColor Green
get-service -ComputerName $xdc -name "Citrix Licensing" |Stop-Service

Write-host "Deleting the $xdc license cache files" -ForegroundColor Green
get-childitem "\\$xdc\$XendriveLetter$path"|Remove-Item

Write-host "Restarting the $xdc Citrix Licensing service" -ForegroundColor Green
get-service -ComputerName $xdc -name "Citrix Licensing" |Start-Service

Write-host "Removing the $xdc CIM session" -ForegroundColor Green
Remove-CimSession $cim
Write-host ""
}
Write-host "Complete" -ForegroundColor Green

}

Function Get-MonthlyUserReport{
<#
 .Synopsis
  Generates a comprehensive report of all Active Directory users with their account status and login history.

 .Description
  This script queries Active Directory for all user accounts and exports detailed information including account names,
  enabled status, creation dates, and last logon information to a CSV file. The script displays a progress bar
  during execution to indicate completion status.
  
  The generated report is saved to C:\Downloads\DCIN-TS_UserReport-Newton.csv and automatically opened upon completion.

 .Example
    Get-MonthlyUserReport
    # Runs the script and generates a complete user report CSV file
   
 .Notes
    Purpose: Audit and documentation of Active Directory user accounts
    Output: CSV file with comprehensive user account information
#>

# Gathering users via requested metrics from Active Directory
Write-Host "Gathering Users from ActiveDirectory..." -ForegroundColor Yellow -BackgroundColor Black -NoNewline 
# Query Active Directory for all users and select only the properties needed for the report
$Users = Get-ADUser -Filter * -Properties * | Select-Object Name, SamAccountName, Enabled, whenCreated, LastLogonTimeStamp, LastLogonDate, CanonicalName
Write-Host "Done." -ForegroundColor Green -BackgroundColor Black

# Initialize the report collection and counter variables
$Report = [System.Collections.Generic.List[System.String]]@()  # Create a generic list to store user information
$total = $Users.Count  # Get total number of users for progress calculation
$cnt = 0  # Initialize counter for progress tracking

# Process each user and add their information to the report
# Display a progress bar showing percentage complete and current user being processed
ForEach($i in $Users){
    $cnt++  # Increment the counter for each user processed
    [int]$percent = $cnt / $total * 100  # Calculate percentage complete
    
    # Create a custom object for each user with relevant properties
    $Report += [pscustomobject] @{ 
        Name = $i.Name;  # Full name of the user
        SamAccountName = $i.SamAccountName;  # Login username
        Enabled = $i.Enabled;  # Account status (enabled/disabled)
        WhenCreated = $i.whenCreated;  # Account creation date
        LastLogonTimeStamp = $i.LastLogonTimeStamp;  # Raw logon timestamp
        LastLogonDate = $i.LastLogonDate;  # Formatted last logon date
        CanonicalName = $i.CanonicalName;  # User's location in AD structure
    }
    
    # Update the progress bar with current status
    Write-Progress -Activity "Generating User Report" -Status $i.SamAccountName -PercentComplete $percent
}

# Export the collected data to a CSV file
# Select specific properties to ensure proper column order in the output file
$Report | 
    Select-Object Name, SamAccountName, Enabled, WhenCreated, LastLogonTimeStamp, LastLogonDate, CanonicalName | 
    Export-CSV C:\Downloads\DCIN-TS_UserReport-Newton.csv -NoTypeInformation -Encoding UTF8 

# Open the CSV file automatically after completion
Invoke-Item C:\Downloads\DCIN-TS_UserReport-Newton.csv
}

Function Get-locked {
Search-ADAccount -LockedOut | ? {$_.enabled -eq $true -and $_.samaccountname -like "*-*"} | select samaccountname
}

Function fix-windowsupdate {
[cmdletbinding ()]
param(
    [parameter (Mandatory = $true)]
    [string]$Target
    )

$target = "n485s40cy"
get-service -ComputerName $target -name wuauserv |Stop-Service
Get-service -ComputerName $target -name  ccmexec | stop-service
get-service -ComputerName $target -name  bits | stop-service
$DateStr = $(get-date).ToString("yyyyMMdd")
$named =  "\\" + $target + "\c$\windows\softwaredistribution\"
$renamed = "softwaredistribution" + $DateStr 
rename-item -path $named -NewName $renamed
get-service -ComputerName $target -name wuauserv |Start-Service
Get-service -ComputerName $target -name  ccmexec | start-service
get-service -ComputerName $target -name  bits | start-service

}

function rename-registrypol {
[cmdletbinding ()]
param(
    [parameter (Mandatory = $true)]
    [string]$Target
    )

if (Test-Connection -ComputerName $target -Count 2 -Quiet) {
$path = "\\" + $target + "\C$\windows\system32\grouppolicy\machine\registry.pol"
$pathold = $path + ".old"
    Try {
        move-item -Path $path -destination $pathold -force  -ErrorAction stop
        write-host "Renaming registry.pol to registry.pol.old successful" -ForegroundColor Green
        Write-Host "Reboot system to recreate registry.pol" -ForegroundColor Yellow
        Write-host ""
        }
    Catch {
        write-host "Renaming Registry.pol to Registry.pol.old failed" -ForegroundColor Red
        Write-host ""
        }
    }
    Else {Write-host "System offline" -ForegroundColor Red ;Write-host ""}
}

Function Fix-Spooler{
[cmdletbinding ()]
param(
    [parameter (Mandatory = $false)]
    [string]$Target
    )

$fpservers = (Get-ADComputer -filter {(OperatingSystem -like "*Windows Server*") -and (Name -like "*FP03") -or (Name -like "*FP07")}).name
$fpservers = $fpservers |sort

if ($target) {$fpservers = $target}

ForEach($fp in $fpservers){
    if (Test-Connection $fp -count 2 -Quiet) {
    Write-host "$fp " -NoNewline
    $Featurestatus = (Get-WmiObject -ComputerName $fp -Query 'select * from win32_serverfeature' -ErrorAction SilentlyContinue | ? {$_.name -like "*print server*"})
    $ServiceStatus = @(Get-Service -ComputerName $fp -Name "Spooler" | Select -property Status,StartType)
    if ($Featurestatus) {$Featurestatus = "Installed"} else {$Featurestatus = "Not Installed"}
    Write-Host "Printer server feature $featurestatus ; Spooler Service $($servicestatus.status) ; $($servicestatus.starttype)" -NoNewline
    if (($Featurestatus -eq "Installed") -and ($($servicstatus.status) -eq "Stopped")) {
        Set-service -ComputerName $fp -Name "Spooler" -StartupType Automatic -Status Running -Confirm:$false
        Write-host "; Print Spooler service set to automatic and started" -NoNewline -ForegroundColor Red
        }
    $Featurestatus = $null
    $ServiceStatus = $null
    Write-Host ""
    }
  }
}

function Whats-up-Doc {
[cmdletbinding ()]
param(
    [parameter (Mandatory = $false)]
    [string]$Target
    )


$serverlist = (Get-ADComputer -Filter {(OperatingSystem -like '*Server*') -and (name -notlike '*XDC*')} -Properties operatingsystem).name

if ($target) {
$target = "N" + $target + "*"
$serverlist = (Get-ADComputer -Filter {(name -like $Target) -and (OperatingSystem -like '*Server*') -and (name -notlike '*XDC*')} -Properties operatingsystem).name
}

Write-host @"
                         /| |\
                        / | | \
                        | | | |     Neeaah, Whats up Doc !?!
                        \ | | /
                         \|w|/    /
                         /_ _\   /     ,
              /\       _:()_():_      /]
              ||_     : ._=Y=_  :    / /
             [)(_\,   ',__\W/ _,'   /  \
             [) \_/\    _/'='\     /-/\)
              [_| \ \  ///  \ '._ / /
              :;   \ \///   / |  '` /
              ;::   \ `|:   : |',_.'
              """    \_|:   : |
                       |:   : |'".
                       /`._.'  \/
                      /  /|   /
                     |  \ /  /
                      '. '. /
                        '. '
                        / \ \
                       / / \'=,
                 .----' /   \ (\__
            snd (((____/     \ \  )
                              '.\_)'
"@

$serverlist = $serverlist | sort
$offline = @()

Foreach ($server in $serverlist) {
        if (Test-Connection $server -Count 2 -Quiet) {
        Write-host "$server is up"
        }
    Else {
    $offline += $server
        Write-host "$server is down" -ForegroundColor Red
        }
}

Write-host ""
Write-host "The Following Systems are down as of $(get-date) $([system.timezoneinfo]::Local.id):" -ForegroundColor Red
$body = "`n Technician: $(whoami)"
$body += "`n Refer to the High Priority Sites Memo on TS Sharepoint Portal for reporting responsibilites."
$body += "`n "
$body += "`n The following systems are down as of $(get-date) $([system.timezoneinfo]::Local.id):"

foreach ($off in $offline) {
    try {
    [string]$canname = (Get-ADComputer $off -properties canonicalname |select canonicalname).canonicalname
    $canname = $canname -replace "newton.pentagon.mil/sites/" , ""
    $canname = $canname -replace "$off" , ""
    write-host $off - $canname -ForegroundColor red
    $body += "`n $off - $canname"
    $canname = $null
    } Catch {Write-host "$off - No Canonical Name in AD" -ForegroundColor Red
            $body += "`n $off = No Canonical Name in AD"
            }
}
$body2 = $body -split '\r?\n' |select-object -skip 5
$wud_folder = "\\n020fp05\admin\Sysadmin Powershell Module\Scripts\Whats-up-Doc\7DayReportHistory\"
$limit = (get-date).AddDays(-7)
get-childitem -path $wud_folder -file |? {$_.CreationTime -lt $limit} |remove-item -force
$last_report = (get-childitem -path $wud_folder|sort name)[-1]
$last_report_body2 = (get-content -path $last_report.FullName)[5..1000]
$new_on_Report = compare-object -ReferenceObject $body2 -DifferenceObject $last_report_body2

$wud_report = "WUD_Report_" + (get-date -Format FileDateTime) + ".txt"
$wud_path = $wud_folder + $wud_report

if (!($target)) {
$body |out-file -FilePath $wud_path
}

if ($($new_on_Report.sideindicator) -like "*=*") {
$body += " `r`n"
$body += " `r`n"
$body += "Changes from previous report: `r`n"
$body += "*************************************** `r`n"
} Else {
$body += " `r`n"
$body += " `r`n"
$body += "Changes from previous report: `r`n"
$body += "*************************************** `r`n"
$body += " `r`n"
$body += "NO CHANGE FROM PREVIOUS REPORT `r`n"
}

Foreach ($new in $new_on_Report) {
if ($new.sideindicator -eq "<=") {$new_status = "*****NEW - NOT ON PREVIOUS WUD REPORT***** `r`n"}
if ($new.sideindicator -eq "=>") { 
    $computernameonly = $($new.inputobject).split("-")[0]
    $computernameonly = $computernameonly.Replace(" ","")
    Try {
    $isInAD = get-adcomputer -identity $computernameonly -ErrorAction SilentlyContinue
        }
        Catch {$isInAD = $null}

    if ($isInAD) {
    $new_status = "*****BACK ONLINE FROM PREVIOUS REPORT***** `r`n"
    } else {$new_status = "*****SERVER DELETED FROM ACTIVE DIRECTORY***** `r`n"}
    $isInAD = $null
}
$body += "$($new.inputobject) $new_status"
}

if (!($Target)) {
$wud_user = "-ncocadmins@newton.pentagon.mil"
    
    #invoke-command -ComputerName N020MBX01.newton.pentagon.mil -ScriptBlock {
    $WUD_toemail = @("-ncocadmins@newton.pentagon.mil")
    # uncomment for testing:
    # $WUD_toemail = @("jfirebaugh@newton.pentagon.mil","jveliz@newton.pentagon.mil")
    Send-MailMessage -SmtpServer relay.newton.pentagon.mil -from $WUD_user -to $WUD_toemail -Body "$body" -Subject "What's Up Doc Report."
    #} 
}

$offline = $null
}

Function Hook-ABrotherUp {
Param(
    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    [string[]]
    $Target
)

Set-PowerCLIConfiguration -DefaultVIServerMode Multiple -confirm:$false | Out-Null
# if testing/runing script in ISE, the stat variable should be reset at the beginning of each run.  
# In powershell console this is not an issue as the variable is forgotten after the script completes.
# it does no harm to remove it at the beginning of the script in either case.
Remove-Variable stat?

# list of valid Vcenter servers for the domain
$vcservers = (
"N066VCSA02",
"N159VCSA02",
"N490VCSA01",
"N447VCSA01")

# Store the Script that runs the queue check in a variable for easy reference in the invoke command 
$script = "net user $Target /active:yes"


Write-host "Authenticating to the Vcenter Servers....." -ForegroundColor Magenta -NoNewline
foreach ($vc in $vcservers){
    if (test-connection -ComputerName $vc -Count 2 -Quiet){
    Write-host "..." -ForegroundColor Magenta -NoNewline
    Connect-VIServer -Server $vc | Out-Null
    }
}

Write-host ""
Write-host ""

if (!($cred)) {
Write-host "Prompting for SMTP server credentials....." -ForegroundColor Magenta
Write-host ""
$global:cred = Get-Credential -message "Enter your NEWTON SMTP server credentials:"
}

Write-host "Unlocking $target on SMTP Servers....." -ForegroundColor Magenta
Write-host ""

Try {
Invoke-VMScript -vm n490smtp01 -ScriptText $script  -GuestCredential $cred -ScriptType bat -erroraction stop | select -ExpandProperty scriptoutput
    }
    catch {
    Write-host "N490SMTP01 Unlock $target Failed.  Check Credentials.  Check SMTP Server Power/Boot State." -ForegroundColor Red
    Write-host ""
    
    }

Try {
Invoke-VMScript -vm n020smtp01 -ScriptText $script  -GuestCredential $cred -ScriptType bat -erroraction stop | select -expandproperty ScriptOutput
    }
    catch {
    Write-host "N020SMTP01 Unlock $target Failed.  Check Credentials.  Check SMTP Server Power/Boot State." -ForegroundColor Red
    Write-host ""
       
    }


write-host "Script complete"


}

function Launch-SMTP{
Connect-VIServer -Server n490vcsa01

 $getvms = (get-vm *smtp* | ? {$_.powerstate -eq "PoweredOn"}) | select name,vmhost

 foreach ($vm in $getvms) {
    if (get-vmhost -name $($vm.vmhost) | where {$_.ConnectionState -eq "disconnected"}) {Set-VMHost -VMHost $($vm.vmhost) -state Connected}
 }

get-vm *smtp* |? {$_.powerstate -eq "PoweredOn"} | Open-VMConsoleWindow

$logpath = "\\newton\admin\SysAdmin Powershell Module\Scripts\BattleRythm\Logs\Log_BattleRhythm.csv"
Add-Content -Path $logpath -Value "`Check-SMTP,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami)" 
}

function install-PowerCLI {
copy "\\n020fp05\admin\VMWare\PowerCLI\VMware-PowerCLI-12.5.0-19195797.zip" "c:\downloads\"
Add-Type -Assembly “system.io.compression.filesystem”
[io.compression.zipfile]::ExtractToDirectory("c:\downloads\VMware-PowerCLI-12.5.0-19195797.zip", "C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules")
Set-PowerCLIConfiguration -scope allusers -invalidCertificateAction ignore -ParticipateInCeip $false -Confirm:$false
#Set-PowerCLIConfiguration -ParticipateInCeip $false -Scope AllUsers  -Confirm:$false  |out-null
#Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Scope AllUsers -Confirm:$false |out-null
Set-PowerCLIConfiguration -DisplayDeprecationWarnings $false -Scope AllUsers -Confirm:$false |out-null
}

function check-smtpQ {
Set-PowerCLIConfiguration -DefaultVIServerMode Multiple -confirm:$false | Out-Null
# if testing/runing script in ISE, the stat variable should be reset at the beginning of each run.  
# In powershell console this is not an issue as the variable is forgotten after the script completes.
# it does no harm to remove it at the beginning of the script in either case.
Remove-Variable stat?

# list of valid Vcenter servers for the domain
$vc = "N490VCSA01"

# Store the Script that runs the queue check in a variable for easy reference in the invoke command 
$script = '(get-childitem -path c:\inetpub\mailroot\queue).count
            [math]::round((get-psdrive c).free /1gb, 2)
            Write-output ""
            get-localuser |? {$_.enabled -eq $true -and $_.name -like "*-admin"} | sort passwordexpires |select name,passwordexpires
            Write-output ""
            Write-output "The following accounts are currently locked out:"
            get-wmiobject -computername $env:computername -class Win32_UserAccount -filter "LocalAccount=True" | ? {($_.name -like "*-admin") -and ($_.lockout -eq $true) } | sort name | select name,lockout'

# Store the script that runs the SMTP Service restart for easy reference in the invoke command if files are piling up in the queue
$restartSMTP = 'get-service smtpsvc |restart-service -force'

Write-host "Authenticating to N490VCSA01....." -ForegroundColor Magenta -NoNewline

    if (test-connection -ComputerName $vc -Count 2 -Quiet){
    Write-host "..." -ForegroundColor Magenta -NoNewline
    Connect-VIServer -Server $vc | Out-Null
    


Write-host ""
Write-host ""

if (!($cred)) {
Write-host "Prompting for SMTP server credentials....." -ForegroundColor Magenta
Write-host ""
$global:cred = Get-Credential -message "Enter your NEWTON SMTP server credentials:"
}

Write-host "Checking SMTP Queues....." -ForegroundColor Magenta
Write-host ""

Try {
$stat1 = Invoke-VMScript -vm n490smtp01 -ScriptText $script  -GuestCredential $cred -ScriptType PowerShell -erroraction stop | select -ExpandProperty scriptoutput |out-string
    }
    catch {
    Write-host "N490SMTP01 SMTP Queue Check Failed.  Check Credentials.  Check SMTP Server Power/Boot State." -ForegroundColor Red
    Write-host ""
    $stat1 = "-1"
    
    }

Try {
$stat2 = Invoke-VMScript -vm n020smtp01 -ScriptText $script  -GuestCredential $cred -ScriptType PowerShell -erroraction stop | select -expandproperty ScriptOutput |out-string
    }
    catch {
    Write-host "N020SMTP01 SMTP Queue Check Failed.  Check Credentials.  Check SMTP Server Power/Boot State." -ForegroundColor Red
    Write-host ""
    $stat2 = "-1"
   
    }

# convert the stat variables from a string to a integer so that the "greater than 5" check will work and run the SMTP service reset if files are piling up in the queue
[int]$Files1 = ($stat1.split("`n"))[0]
[int]$Files2 = ($stat2.split("`n"))[0]

$free1 = ($stat1.split("`n"))[1]
$free2 = ($stat2.split("`n"))[1]


write-host "There are $Files1 files in the queue on N490SMTP01.  Free Space $free1 GB"
write-host "There are $Files2 files in the queue on N020SMTP01.  Free Space $free2 GB"
Write-host ""

$accountinfo1 = $stat1.split("`n")[4..($stat1.split("`n").count - 4)] -join "`n"
$accountinfo2 = $stat2.split("`n")[4..($stat2.split("`n").count - 4)] -join "`n"

Write-host "Accounts Status on N490SMTP01:"
Write-host $accountinfo1
Write-host "Accounts Status on N020SMTP01:"
Write-host $accountinfo2

$body = "There are $Files1 files in the queue on N490SMTP01.  Free Space $free1 GB `n`r"
$body += "There are $Files2 files in the queue on N020SMTP01.  Free Space $free2 GB `n`r"
$body += "`n`r"
$body += "Accounts Status on N490SMTP01: `n"
$body += $accountinfo1 + "`n"
$body += "Accounts Status on N020SMTP01: `n"
$body += $accountinfo2 + "`n"

$Que_user = "-ncocadmins@newton.pentagon.mil"
    
    #invoke-command -ComputerName N020MBX02.newton.pentagon.mil -ScriptBlock {
    $Que_toemail = @("-ncocadmins@newton.pentagon.mil")
    # uncomment for testing:
    # $WUD_toemail = @("jfirebaugh@newton.pentagon.mil","jveliz@newton.pentagon.mil")
    Send-MailMessage -SmtpServer relay.newton.pentagon.mil -from $Que_user -to $Que_toemail -Body "$body" -Subject "SMTP Queue Check Report."
    #} 

if (($stat1 -eq -1) -or ($stat2 -eq -1)) { $global:cred = $null }

if ($Files1 -gt 5) { 

do{
    $prompt = Read-Host -Prompt “Restart SMTP service on N490SMTP01? Y/N?”
    If($prompt -eq “y” -or $prompt -ne “n”){
        break
    }else{
        Write-Host “invalid response, please try again” -ForegroundColor Red
    }
}
until($prompt -eq “y” -or $prompt -eq “n”)

If($prompt -eq “y”){get-vm n490smtp01 | Invoke-VMScript -ScriptText $restartSMTP -GuestCredential $cred -ScriptType PowerShell}
else{    }

 }

if ($Files2 -gt 5) { 

do{
    $prompt = Read-Host -Prompt “Restart SMTP service on N020SMTP01? Y/N?”
    If($prompt -eq “y” -or $prompt -ne “n”){
        break
    }else{
        Write-Host “invalid response, please try again” -ForegroundColor Red
    }
}
until($prompt -eq “y” -or $prompt -eq “n”)

If($prompt -eq “y”){get-vm n020smtp01 | Invoke-VMScript -ScriptText $restartSMTP -GuestCredential $cred -ScriptType PowerShell}
else{    }

 }

Write-host "Checks complete." (get-date) -ForegroundColor Magenta
Write-host ""

}Else{Write-Host "Cannot connect to N490VCSA01..... Please check vcenter server!"}
Write-host "Pausing for 4 hours.  Will run again at $((get-date).addhours(4)).  To cancel next run close this window or press CTRL-C." -ForegroundColor Green
start-sleep -Seconds 14400
check-smtpQ
}

Function Check-SvcLastPassSet {
# SVC accounts with passwords more than 350 days old (2 weeks before the 1 year mark to facilitate password change coordination with the service POC).
# will need to capture the results to an array.  If PasswordLastSet is null, that means the password has never been reset since the account was created.  
# In that case, password age will be whenCreated 

$Folder = "\\newton\admin\SysAdmin Powershell Module\Scripts\Check-SvcLastPassSet\Logs"
$SvcAccts = Get-ADUser -Filter {(SamAccountName -like "*newton-sysadmin*") -or (SamAccountName -like "*SVC*") -and (Enabled -eq $True)} -Properties PasswordLastSet,whenCreated
$Yr = Get-Date -Format MM-yyyy


# - Go through each service account and if there was no date for a password last set, it will be changed to the date the account was created.
ForEach($fp in $SvcAccts){ If($fp.PasswordLastSet -eq $null){ $fp.PasswordLastSet = $s.WhenCreated } }

# - Go through each service account and add each one that hasn't been changed in 350 days or more to the list of bad ones.
ForEach($fp in $SvcAccts){
    If($fp.PasswordLastSet -lt (Get-Date).adddays(-350)){ 
        $Report = [pscustomobject]@{ SamAccountName = $fp.SamAccountName ; WhenCreated = $fp.whenCreated ; PasswordLastSet = $fp.PasswordLastSet }
        $Report | Select SamAccountName,WhenCreated,PasswordLastSet | Export-Csv -Path $Folder\ExpiredPassword-Svc-$Yr.csv -NoType -Encoding UTF8 -Append
        }
    }


# - Checks if there were service accounts found and display the result.
If(!(Test-Path -Path $Folder\ExpiredPassword-Svc-$Yr.csv -ErrorA Si)){Write-Host "No Accounts in violation!" -f Green}
Else{Write-Host "Found Accounts in Violation, Loading CSV File..." -f Yellow -b Black ; Invoke-Item -Path $Folder\ExpiredPassword-Svc-$Yr.csv}

$logpath = "\\newton\admin\SysAdmin Powershell Module\Scripts\BattleRythm\Logs\Log_BattleRhythm.csv"
Add-Content -Path $logpath -Value "`Check-SvcLastPassSet,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami)" 

}

Function Get-TestGrpApps {
 param (
    [Parameter(ValueFromPipeline=$true)]
    [string]$NameRegex = ''
)
$ComputerName = @(Import-Csv -Path "\\newton\admin\SysAdmin Powershell Module\Scripts\Update-SysAdmin\SysAdminComputers.csv")
    foreach ($comp in $ComputerName) {
    if (Test-Connection $comp -count 2 -quiet) {
            $loggedonuser = @(wmic /node:$comp computersystem get username)
            $loggedonuser = $loggedonuser.Split('',[System.StringSplitOptions]::RemoveEmptyEntries)
            $keys = '','\Wow6432Node'
            foreach ($key in $keys) {
                try {
                    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $comp)
                    $apps = $reg.OpenSubKey("SOFTWARE$key\Microsoft\Windows\CurrentVersion\Uninstall").GetSubKeyNames()
                } catch {
                    continue
                }

                foreach ($app in $apps) {
                    $program = $reg.OpenSubKey("SOFTWARE$key\Microsoft\Windows\CurrentVersion\Uninstall\$app")
                    $name = $program.GetValue('DisplayName')
                    if ($name -and $name -match $NameRegex) {
                        [pscustomobject]@{
                            ComputerName = $comp
                            DisplayName = $name
                            DisplayVersion = $program.GetValue('DisplayVersion')
                            Publisher = $program.GetValue('Publisher')
                            InstallDate = $program.GetValue('InstallDate')
                            UninstallString = $program.GetValue('UninstallString')
                            Bits = $(if ($key -eq '\Wow6432Node') {'64'} else {'32'})
                            Path = $program.name
                            User = $loggedonuser -join " - "
                        }
                    }
                }
            }
        }
    }
}

function GetBL {
$LapsLockerInfo.text = " "
Try
{
Get-adcomputer $($compname.text) -erroraction silentlycontinue
$compdn = (get-adcomputer $($compname.text)).DistinguishedName
$RawdataHash = (Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase $compdn -Properties 'msFVE-RecoveryPassword') | Select name,msFVE-RecoveryPassword
if ($RawdataHash -eq $null) {$LapsLockerInfo.text = "No Bitlocker Recovery Data in AD."}
    Foreach ($R in $RawdataHash) {
    $LapsLockerInfo.text += "`r`n" + $R.name
    $LapsLockerInfo.text += "`r`n" + $R."msFVE-RecoveryPassword"
    $LapsLockerInfo.text += "`r`n"
    }
}
catch
{
$LapsLockerInfo.text = "Computer Name not found `r`nOr access denied to the msFVE-RecoveryPassword attribute."
}

}

Function LapQuery {
Try
{
Get-adcomputer $($compname.text) -erroraction silentlycontinue

$lapsstring = ((get-adcomputer $($compname.text) -properties ms-Mcs-AdmPwd)."ms-Mcs-AdmPwd") 

$LapsLockerInfo.text = $lapsstring
$LapsLockerInfo.text += "`r`nPhonetic Spelling"
$LapsLockerInfo.text += "`r`n------------"

$LapsArray = @($lapsstring.ToCharArray() | %{[char]$_})

foreach ($L in $LapsArray) {
$AscHex = '{0:x2}' -f ([int][char]$L)
$HexArray += @($AscHex)
}

$CharArray = @(
[PSCustomObject]@{inchar= "61";  Spellsound = "a - Lower case a as in alfa"}
[PSCustomObject]@{inchar= "62";  Spellsound = "b - Lower case b as in bravo"}
[PSCustomObject]@{inchar= "63";  Spellsound = "c - Lower case c as in charlie"}
[PSCustomObject]@{inchar= "64";  Spellsound = "d - Lower case d as in delta"}
[PSCustomObject]@{inchar= "65";  Spellsound = "e - Lower case e as in echo"}
[PSCustomObject]@{inchar= "66";  Spellsound = "f - Lower case f as in foxtrot"}
[PSCustomObject]@{inchar= "67";  Spellsound = "g - Lower case g as in golf"}
[PSCustomObject]@{inchar= "68";  Spellsound = "h - Lower case h as in hotel"}
[PSCustomObject]@{inchar= "69";  Spellsound = "i - Lower case i as in india"}
[PSCustomObject]@{inchar= "6a";  Spellsound = "j - Lower case j as in juliett"}
[PSCustomObject]@{inchar= "6b";  Spellsound = "k - Lower case k as in kilo"}
[PSCustomObject]@{inchar= "6c";  Spellsound = "l - Lower case l as in lima"}
[PSCustomObject]@{inchar= "6d";  Spellsound = "m - Lower case m as in mike"}
[PSCustomObject]@{inchar= "6e";  Spellsound = "n - Lower case n as in november"}
[PSCustomObject]@{inchar= "6f";  Spellsound = "o - Lower case o as in oscar"}
[PSCustomObject]@{inchar= "70";  Spellsound = "p - Lower case p as in papa"}
[PSCustomObject]@{inchar= "71";  Spellsound = "q - Lower case q as in quebec"}
[PSCustomObject]@{inchar= "72";  Spellsound = "r - Lower case r as in romeo"}
[PSCustomObject]@{inchar= "73";  Spellsound = "s - Lower case s as in sierra"}
[PSCustomObject]@{inchar= "74";  Spellsound = "t - Lower case t as in tango"}
[PSCustomObject]@{inchar= "75";  Spellsound = "u - Lower case u as in uniform"}
[PSCustomObject]@{inchar= "76";  Spellsound = "v - Lower case v as in victor"}
[PSCustomObject]@{inchar= "77";  Spellsound = "w - Lower case w as in whiskey"}
[PSCustomObject]@{inchar= "78";  Spellsound = "x - Lower case x as in xray"}
[PSCustomObject]@{inchar= "79";  Spellsound = "y - Lower case y as in yankee"}
[PSCustomObject]@{inchar= "7a";  Spellsound = "z - Lower case z as in zulu"}
[PSCustomObject]@{inchar= "41";  Spellsound = "A - Upper case A as in Alfa"}
[PSCustomObject]@{inchar= "42";  Spellsound = "B - Upper case B as in Bravo"}
[PSCustomObject]@{inchar= "43";  Spellsound = "C - Upper case C as in Charlie"}
[PSCustomObject]@{inchar= "44";  Spellsound = "D - Upper case D as in Delta"}
[PSCustomObject]@{inchar= "45";  Spellsound = "E - Upper case E as in Echo"}
[PSCustomObject]@{inchar= "46";  Spellsound = "F - Upper case F as in Foxtrot"}
[PSCustomObject]@{inchar= "47";  Spellsound = "G - Upper case G as in Golf"}
[PSCustomObject]@{inchar= "48";  Spellsound = "H - Upper case H as in Hotel"}
[PSCustomObject]@{inchar= "49";  Spellsound = "I - Upper case I as in India"}
[PSCustomObject]@{inchar= "4a";  Spellsound = "J - Upper case J as in Juliett"}
[PSCustomObject]@{inchar= "4b";  Spellsound = "K - Upper case K as in Kilo"}
[PSCustomObject]@{inchar= "4c";  Spellsound = "L - Upper case L as in Lima"}
[PSCustomObject]@{inchar= "4d";  Spellsound = "M - Upper case M as in Mike"}
[PSCustomObject]@{inchar= "4e";  Spellsound = "N - Upper case N as in November"}
[PSCustomObject]@{inchar= "4f";  Spellsound = "O - Upper case O as in Oscar"}
[PSCustomObject]@{inchar= "50";  Spellsound = "P - Upper case P as in Papa"}
[PSCustomObject]@{inchar= "51";  Spellsound = "Q - Upper case Q as in Quebec"}
[PSCustomObject]@{inchar= "52";  Spellsound = "R - Upper case R as in Romeo"}
[PSCustomObject]@{inchar= "53";  Spellsound = "S - Upper case S as in Sierra"}
[PSCustomObject]@{inchar= "54";  Spellsound = "T - Upper case T as in Tango"}
[PSCustomObject]@{inchar= "55";  Spellsound = "U - Upper case U as in Uniform"}
[PSCustomObject]@{inchar= "56";  Spellsound = "V - Upper case V as in Victor"}
[PSCustomObject]@{inchar= "57";  Spellsound = "W - Upper case W as in Whiskey"}
[PSCustomObject]@{inchar= "58";  Spellsound = "X - Upper case X as in Xray"}
[PSCustomObject]@{inchar= "59";  Spellsound = "Y - Upper case Y as in Yankee"}
[PSCustomObject]@{inchar= "5a";  Spellsound = "Z - Upper case Z as in Zulu"} 
[PSCustomObject]@{inchar= "30";  Spellsound = "0 - Number Zero"}
[PSCustomObject]@{inchar= "31";  Spellsound = "1 - Number One"}
[PSCustomObject]@{inchar= "32";  Spellsound = "2 - Number Two"}
[PSCustomObject]@{inchar= "33";  Spellsound = "3 - Number Three"}
[PSCustomObject]@{inchar= "34";  Spellsound = "4 - Number Four"}
[PSCustomObject]@{inchar= "35";  Spellsound = "5 - Number Five"}
[PSCustomObject]@{inchar= "36";  Spellsound = "6 - Number Six"}
[PSCustomObject]@{inchar= "37";  Spellsound = "7 - Number Seven"}
[PSCustomObject]@{inchar= "38";  Spellsound = "8 - Number Eight"}
[PSCustomObject]@{inchar= "39";  Spellsound = "9 - Number Nine"} 
[PSCustomObject]@{inchar= "60";  Spellsound = "` - 'Back Tick' or 'Back Quote' or 'Grave Accent'"}
[PSCustomObject]@{inchar= "7e";  Spellsound = "~ - Tilde"}
[PSCustomObject]@{inchar= "21";  Spellsound = "! - Exclamation Mark"}
[PSCustomObject]@{inchar= "40";  Spellsound = "@ - At Sign"}
[PSCustomObject]@{inchar= "23";  Spellsound = "# - Pound Sign"}
[PSCustomObject]@{inchar= "24";  Spellsound = "`$ - Dollar Sign"}
[PSCustomObject]@{inchar= "25";  Spellsound = "% - Percent"}
[PSCustomObject]@{inchar= "5e";  Spellsound = "^ - Carat"}
[PSCustomObject]@{inchar= "26";  Spellsound = "& - Ampersand"}
[PSCustomObject]@{inchar= "2a";  Spellsound = "* - Asterisk"}
[PSCustomObject]@{inchar= "28";  Spellsound = "( - Open Parentheses"}
[PSCustomObject]@{inchar= "29";  Spellsound = ") - Close Parentheses"}
[PSCustomObject]@{inchar= "2d";  Spellsound = "- - Dash"}
[PSCustomObject]@{inchar= "5f";  Spellsound = "_ - Underscore"}
[PSCustomObject]@{inchar= "3d";  Spellsound = "= - Equal Sign"}
[PSCustomObject]@{inchar= "2b";  Spellsound = "+ - Plus Sign"}
[PSCustomObject]@{inchar= "7b";  Spellsound = "{ - Open Curly Brace"}
[PSCustomObject]@{inchar= "7d";  Spellsound = "} - Close Curly Brace"}
[PSCustomObject]@{inchar= "5b";  Spellsound = "[ - Open Bracket"}
[PSCustomObject]@{inchar= "5d";  Spellsound = "] - Close Bracket"}
[PSCustomObject]@{inchar= "5c";  Spellsound = "\ - Back Slash"}
[PSCustomObject]@{inchar= "7c";  Spellsound = "| - Pipe Symbol"}
[PSCustomObject]@{inchar= "3b";  Spellsound = "; - Semicolon"}
[PSCustomObject]@{inchar= "3a";  Spellsound = ": - Colon"}
[PSCustomObject]@{inchar= "27";  Spellsound = "' - 'Apostrophe' or 'Single Quote'"}
[PSCustomObject]@{inchar= "22";  Spellsound = "`" - Double Quote"}
[PSCustomObject]@{inchar= "2c";  Spellsound = ", - Comma"}
[PSCustomObject]@{inchar= "2e";  Spellsound = ". - Period"}
[PSCustomObject]@{inchar= "2f";  Spellsound = "/ - Forward Slash"}
[PSCustomObject]@{inchar= "3c";  Spellsound = "< - 'Less than' or 'Open Angle Bracket'"}
[PSCustomObject]@{inchar= "3e";  Spellsound = "> - 'Greater than' or 'Close Angle Bracket'"}
[PSCustomObject]@{inchar= "3f";  Spellsound = "? - Question Mark"}
[PSCustomObject]@{inchar= "20";  Spellsound = "  - Space"}
)
foreach ($H in $HexArray) {
$LapsLockerInfo.text += "`r`n" + ($CharArray -match $H).spellsound

}

}
catch
{
$LapsLockerInfo.text = "Computer Name not found or access denied `r`nto the ms-Mcs-AdmPwd attribute"
}

}

Function LapsLocker {
$HexArray = $null
$lapsstring = "TryAgain"
$LapsArray = $null

# adds powershell assemblies to support form creation.
#
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


$form = New-Object System.Windows.Forms.Form
$form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
$form.Text = 'Laps-Locker Tool'
$form.Size = New-Object System.Drawing.Size(600,600)
$form.StartPosition = 'CenterScreen'

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(10,520)
$OKButton.Size = New-Object System.Drawing.Size(100,23)
$OKButton.Text = "OK, I'm done"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$OKButton.Enabled = $true
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$GetLapsButton = New-Object System.Windows.Forms.Button
$GetLapsButton.Location = New-Object System.Drawing.Point(140,20)
$GetLapsButton.Size = New-Object System.Drawing.Size(100,23)
$GetLapsButton.Text = ".\LclAdmin"
$GetLapsButton.Add_Click({LapQuery})
$GetLapsButton.Enabled = $true
$form.AcceptButton = $GetLapsButton
$form.Controls.Add($GetLapsButton)

$GetBLButton = New-Object System.Windows.Forms.Button
$GetBLButton.Location = New-Object System.Drawing.Point(140,50)
$GetBLButton.Size = New-Object System.Drawing.Size(100,23)
$GetBLButton.Text = "Get BitLocker"
$GetBLButton.Add_Click({GetBL})
$GetBLButton.Enabled = $true
$form.AcceptButton = $GetBLButton
$form.Controls.Add($GetBLButton)

# Creates the first form field text and places it on the form 10 pixels from the edge and 20 pixels from the top.
# 
$CompNameLabel = New-Object System.Windows.Forms.Label
$CompNameLabel.Location = New-Object System.Drawing.Point(10,20)
$CompNameLabel.Size = New-Object System.Drawing.Size(720,20)
$CompNameLabel.Text = 'Computer Name:'
$form.Controls.Add($CompNameLabel)

# Creates the first form field box and places it on the form 10 pixels from the edge and 40 pixels from the top.
# 
$CompName = New-Object System.Windows.Forms.TextBox
$CompName.Location = New-Object System.Drawing.Point(10,40)
$CompName.Size = New-Object System.Drawing.Size(120,20)
$form.Controls.Add($CompName)

# Creates the first form field text and places it on the form 10 pixels from the edge and 20 pixels from the top.
# 
$LapsLockerInfoLabel = New-Object System.Windows.Forms.Label
$LapsLockerInfoLabel.Location = New-Object System.Drawing.Point(10,80)
$LapsLockerInfoLabel.Size = New-Object System.Drawing.Size(720,20)
$LapsLockerInfoLabel.Text = '.\LclAdmin Password:'
$form.Controls.Add($LapsLockerInfoLabel)

# Creates the first form field box and places it on the form 10 pixels from the edge and 40 pixels from the top.
# 
$LapsLockerInfo = New-Object System.Windows.Forms.TextBox
$LapsLockerInfo.Location = New-Object System.Drawing.Point(10,100)
$LapsLockerInfo.Size = New-Object System.Drawing.Size(560,400)
$LapsLockerInfo.Multiline = $true
$LapsLockerInfo.WordWrap = $true
$LapsLockerInfo.font = "Courier,10"
$LapsLockerInfo.wordwrap = $false
$form.Controls.Add($LapsLockerInfo)

# sets the form to appear on top of other windows even if the user clicks off of the form.
#
$form.Topmost = $true

# sets textbox1 as the starting place for filling out the form.
#
$form.Add_Shown({$CompName.Select()})

# collects the results of the form
$result = $form.ShowDialog()


# **********FORM ENDS*************
# **********FORM ENDS*************
# **********FORM ENDS*************
# **********FORM ENDS*************
# **********FORM ENDS*************
# **********FORM ENDS*************


}

function Get-loggedon {

Param(
    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    [string[]]
    $glocomputer
)

    if (Test-Connection $glocomputer -Count 2 -Quiet) {
        try {
            $user = $null
            $user = gwmi -Class win32_computersystem -ComputerName $glocomputer | select -ExpandProperty username -ErrorAction Stop
        } catch { "Not logged on"; return }
        try {
            Write-Host "WinRM : " -f DarkGray -no 
            $WinRM = Set-Service -ComputerName $glocomputer -Name WinRM -Status Running  -ErrorAction Si
            If($WinRM -eq $false){Write-Host "Failed to Start." -f Red} Else{ Write-Host "Successfully Started." -f Green}
            Write-Host "WManSvc : " -f DarkGray -no 
            $WManSvc = Set-Service -ComputerName $glocomputer -Name WManSvc -Status Running  -ErrorAction Si
            If($WManSvc -eq $false){Write-Host "Failed to Start." -f Red} Else{ Write-Host "Successfully Started." -f Green}
            if ((Invoke-Command -ComputerName $glocomputer -ErrorAction Stop -ScriptBlock { Get-Process logonui }) -and ($user)) {
                "Workstation locked by $user"
            } else {
                "Verifying with qwinsta"
                qwinsta /server:$glocomputer
            }
        } catch { if ($user) { "$user logged on" } }
    } else { "$glocomputer Offline" }
}

Function Fix-SCCMClient { 
Param(
    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    [string[]]
    $Target
)
 
If(Test-Connection $Target -count 2 -Quiet -ErrorA Si)
{  
Write-Host "WinRM : " -f DarkGray -no 
$WinRM = Set-Service -ComputerName $Target -Name WinRM -Status Running  -ErrorAction Si
If($WinRM -eq $false){Write-Host "Failed to Start." -f Red} Else{ Write-Host "Successfully Started." -f Green}
Write-Host "WManSvc : " -f DarkGray -no 
$WManSvc = Set-Service -ComputerName $Target -Name WManSvc -Status Running  -ErrorAction Si
If($WManSvc -eq $false){Write-Host "Failed to Start." -f Red} Else{ Write-Host "Successfully Started." -f Green}


    Invoke-Command -ComputerName $Target -ScriptBlock {
    $sms = new-object –comobject “Microsoft.SMS.Client”
    if ($sms.GetAssignedSite() –ne “N24”) { $sms.SetAssignedSite(“N24”) }
    Write-host "SCCM Site set to 'N24'" -f Green
    #Clear the configuration manager cache
    ## Initialize the CCM resource manager com object
    [__comobject]$CCMComObject = New-Object -ComObject 'UIResource.UIResourceMgr'
    ## Get the CacheElementIDs to delete
    $CacheInfo = $CCMComObject.GetCacheInfo().GetCacheElements()
    ## Remove cache items
    ForEach ($CacheItem in $CacheInfo) {
    $null = $CCMComObject.GetCacheInfo().DeleteCacheElement([string]$($CacheItem.CacheElementID))
}
    Remove-Item -Path ‘HKLM:\SOFTWARE\Microsoft\SystemCertificates\SMS\Certificates\*’ -force -ErrorAction si
    restart-service ccmexec -Force -ErrorAction si
    Write-host "SCCM Client certificate deleted and CCMExec service restarted to force reinstall of client certificate." -f Green
    Start-sleep 20
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000001}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000002}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000003}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000010}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000011}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000012}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000021}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000022}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000023}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000024}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000025}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000026}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000027}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000031}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000032}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000037}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000040}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000041}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000042}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000043}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000051}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000061}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000062}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000063}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000101}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000102}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000103}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000104}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000105}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000106}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000107}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000108}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000109}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000111}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000112}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000113}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000114}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000115}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000116}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000121}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000122}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000123}" -ErrorAction si | out-null
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000131}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000221}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000222}" -ErrorAction si | out-null
 #  Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000223}" -ErrorAction si | out-null
    Write-host "All SCCM Actions Initiated" -f green
    }

}
Else
{
Write-host "FIX ACTION FAILED - $($Target.toupper()) OFFLINE               " -ForegroundColor Red -BackgroundColor black 
}
}

Function Check-DriveSpace {

[cmdletbinding()]
param([Parameter(Mandatory=$false)][switch]$ShowOutput)

Function brk {Write-Host ""}

Function Show-Output {
    param([switch]$Low,[switch]$Critical,[switch]$AllGood,[switch]$Tier3,[switch]$FailedPing)
    
    If($Low){
        Write-Host "$outputname - " -f Yellow -no ; Write-Host "Capacity : " -f Gray -no ; Write-Host "$Size" -f Yellow -no ; Write-Host "GB " -f Gray -no ; Write-Host " | " -f DarkGray -no 
        Write-Host "Free Space : " -f Gray -no ; Write-Host "$sFree" -f Yellow -no ; Write-Host "GB " -f Gray -no ; Write-Host " | " -f DarkGray -no 
        Write-Host "Percent Free : " -f Gray -no ; Write-Host "$pfree " -f Yellow -no ; Write-Host " | " -f DarkGray -no 
        Write-Host "Status : " -f Gray -no ; Write-Host "LOW" -f Yellow -b Black
        }
    If($Critical){
        Write-Host "$outputname - " -f Yellow -no ; Write-Host "Capacity : " -f Gray -no ; Write-Host "$Size" -f Yellow -no ; Write-Host "GB " -f Gray -no ; Write-Host " | " -f DarkGray -no 
        Write-Host "Free Space : " -f Gray -no ; Write-Host "$sFree" -f Yellow -no ; Write-Host "GB " -f Gray -no ; Write-Host " | " -f DarkGray -no 
        Write-Host "Percent Free : " -f Gray -no ; Write-Host "$pfree " -f Yellow -no ; Write-Host " | " -f DarkGray -no 
        Write-Host "Status : " -f Gray -no ; Write-Host "CRITICAL" -f Red -b Black
        }
    If($AllGood){Write-Host "---- ALL DRIVES ABOVE THRESHOLD ON " -f Green -no ; Write-Host "$Svr " -f Cyan -no ; Write-Host "----" -f Green}
    If($Tier3){Write-Host "---- TIER 3 REQUIRED FOR " -f Magenta -no ; Write-Host "$Svr " -f Cyan -no ; Write-Host "----" -f Magenta}
    If($FailedPing){Write-Host "---- CONNECTION FAILED FOR " -f Magenta -no ; Write-Host "$Svr " -f Cyan -no ; Write-Host "----" -f Magenta}

}

$Servers = Get-ADComputer -Filter {(Name -like "*sol*") -and (name -like "*sql*") -and (Name -like "*DC0*") -and (Name -notlike "*XDC*") -or (Name -like "*MECM*") -or (Name -like "*WSUS*") -or (Name -like "*FP0*") -or (Name -like "*MBX0*") -and (OperatingSystem -like "Windows Server*")} | Sort Name | Select -expand Name
$ReportDate = Get-Date -f "dd MMM yyyy"
$tscnt = $Servers.Count
$scnt = 1

Write-Host "CHECKING DRIVE SPACE OF ALL SERVERS..." -f Yellow -b Black
If($ShowOutput){
    Write-Host "THE FOLLOWING SERVERS HAVE EITHER " -f White -b Black -no ; Write-Host "LOW " -f Yellow -b Black -no ; Write-Host "OR " -f White -b Black -no ; Write-Host "CRITICAL " -f Red -b Black -no ; 
    Write-Host "SPACE AVAILABLE..." -f White -b Black ; brk
}

$htmlhead = "
    <html>
    <style>
    BODY{font-family: Arial; font-size: 8pt;}
    H1{font-size: 16px;}
    H2{font-size: 14px;}
    H3{font-size: 12px;}
    TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
    TH{border: 1px solid black; font-size: 12pt; background: #dddddd; padding: 8px; color: #000000;}
    TD{border: 1px solid black; font-size: 10pt; padding: 8px;}
    td.pass{background: #7FFF00;}
    td.warn{background: #FFE600;}
    td.fail{background: #FF0000; color: #ffffff;}
    td.info{background: #85D4FF;}
    </style>
    <body>
    <h1 align='center'>Server Disk Space Report</h1>
    <h3 align='center'> Generated: $ReportDate</h3>"

$htmlTable = "
    <table align='center'>
    <tr>
    <th><b>Server</b></th>
    <th><b>Drive Letter</b></th>
    <th><b>Label</b></th>
    <th><b>Disk Capacity (GB)</b></th>
    <th><b>Free Space (GB)</b></th>
    <th><b>Percentage Available</b></th>
    </tr>"

ForEach($Svr in $Servers){
    
    # Writing the Progress
    [int]$Percent = $scnt / $tscnt * 100
    Write-Progress -Activity "CHECKING DRIVES OF $Svr" -Status "SERVER $scnt OF $tscnt | $Percent%" -PercentComplete $Percent
    $scnt++

    If($ShowOutput){Write-Host "------------ PROCESSING $Svr ------------" -f White}

    $Ping = Test-Connection $Svr -Count 2 -ErrorAction Si 
    If($Ping){
        Try{$Drives = Get-WmiObject -Class Win32_Volume -ComputerName $Svr | Select SystemName,DriveLetter,Label,Capacity,FreeSpace
        $drivesgood = $true

        # Go through each drive
        ForEach($Dr in $Drives){
            $Letter = $Dr.DriveLetter
            $Name = $Dr.Label
            If($Letter.Length -gt 0){$outputname = "$Letter $Name"}Else{$outputname = "$Name"}
            [int]$Size = [math]::Round(([long]$Dr.Capacity / 1GB),2)
            [int]$sFree = [math]::Round(([long]$Dr.FreeSpace / 1GB),2)
            If($sFree -gt 0){[int]$pFree = 100 * ([math]::Round($sFree / $Size,2))} Else{$pFree = 0}

            Switch($Letter){
                'A:'{ }
                'Z:'{ }
                Default{
                    If($Dr.Label -eq 'BOOT' -or ($dr.Label -eq $null -or ($dr.Capacity -lt 1 -or ($dr.Label -like "*VMware Tools*")))){ } Else{
                    If($pFree -gt 20){ } # Drive is over 20% Free
                    ElseIf($pFree -lt 11){ # Drive is under 11% Free
                        $htmlTable += "<tr>
                                       <td>$Svr</td>
                                       <td>$Letter</td>
                                       <td>$Name</td>
                                       <td>$Size GB</td>
                                       <td>$sFree GB</td>
                                       <td class='fail'>$pFree%</td>
                                       </tr>"
                        $drivesgood = $false ; If($ShowOutput){Show-Output -Critical}
                    }
                    Else{ # Drive is between 11-20 Percent Free
                       # $htmlTable += "<tr>
                       #                <td>$Svr</td>
                       #                <td>$Letter</td>
                       #                <td>$Name</td>
                       #                <td>$Size GB</td>
                       #                <td>$sFree GB</td>
                       #                <td class='warn'>$pFree%</td>
                       #                </tr>"
                       # $drivesgood = $false ; If($ShowOutput){Show-Output -Low}
                    }
                    }
                } # End of Default
            } # End of Switch
        } # End of Each Drive
        If($drivesgood -eq $true){If($ShowOutput){Show-Output -AllGood}}
    }Catch{If($ShowOutput){Show-Output -Tier3}}
    }Else{If($ShowOutput){Show-Output -FailedPing}}
    If($ShowOutput){brk ; brk}
} # End of ForEach Server

$htmlTable += "</table>"
$htmltail = "</body></html>"
$htmlreport = $htmlhead + $htmlTable + $htmltail
$User = $env:USERNAME.Replace("-admin","")
$smtpsettings = @{
    To = "-PMOSystemEngineers@newton.pentagon.mil"
    Cc = "-nocsustainment@newton.pentagon.mil"
    From = "$User@newton.pentagon.mil"
    Subject = "Server Disk Space Report"
    SmtpServer = "relay.newton.pentagon.mil"
    }
#Send-MailMessage @smtpsettings -Body $htmlreport -BodyAsHtml 
#invoke-command -ComputerName N020MBX01.newton.pentagon.mil -ScriptBlock {
    Send-MailMessage @smtpsettings -Body $htmlreport -BodyAsHtml
#}   
$logpath = "\\newton\admin\SysAdmin Powershell Module\Scripts\BattleRythm\Logs\Log_BattleRhythm.csv"
Add-Content -Path $logpath -Value "`Check-DriveSpace,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami)"                 
}

Function Mirror-Permissions {


cls
$gcnt = 1
$gtcnt = 1
$options = @('y','n')
$Tier3 = [System.Collections.Generic.List[System.String]]@()
$Report = @()
$Folder = "\\newton\admin\Scripts\Permission Cloner" 
Function brk{Write-Host ""}


Write-Host "Welcome to Permission Cloner." -f Cyan ; Write-Host "---------------------------------" -f Gray
Write-Host "INSTRUCTIONS...
- The First account you enter is the account that has the permissions you are wanting to use as your cloner.
- The Second account you enter is the account that the permissions will be added to.
- Any permissions requiring Tier3 will be noted, PAY ATTENTION..
- Process will output results per security group." -f Yellow -b Black ; brk

# Asking for First Account
Do {
    $Cloner = Read-Host "Please Enter the Username of the Account you are CLONING FROM" ; 
    Try { $testacc = Get-ADUser $Cloner ; Write-Host "Account Found!" -f Green ; $chk = $true ; brk}
    Catch { Write-Host "Account NOT Found!" -f Red ; $chk = $false}
    }Until($chk -eq $true)

# Outputting Account Name and Security Groups
$accname1 = Get-ADUser -Identity $Cloner | Select -expand Name
Write-Host "Account Chosen : " -f Yellow -no ; Write-Host $accname1 -f Magenta
$Permissions = (Get-ADUser -Identity $Cloner).MemberOf
Write-Host "Total Groups to Clone : " -f Yellow -no ; Write-Host $Permissions.Count -f Magenta

#Opening Space
brk ; brk ; brk

# Asking for Second Account
Do {
    $Clonee = Read-Host "Please Enter the Username of the Account you are CLONING TO" ; 
    Try { $testacc = Get-ADUser $Clonee ; Write-Host "Account Found!" -f Green ; $chk = $true ; brk}
    Catch { Write-Host "Account NOT Found!" -f Red ; $chk = $False}
    }Until($chk -eq $true)

# Gathering account name
$accname2 = Get-ADUser -Identity $Clonee | Select -expand Name
Write-Host "Account Chosen : " -f Yellow -no ; Write-Host $accname2 -f Magenta

# Opening Space
brk ; brk ; brk

# Notice of process beginning with opt out
Write-Host "!!! ATTENTION !!!" -f Red -b Black
Write-Host "You are about to clone the permissions of " -f Yellow -no ; Write-Host $accname1 -f Magenta -no ; Write-Host " to the account " -f Yellow -no ; Write-Host $accname2 -f Magenta -no ; Write-Host "..." -f Gray
Write-Host "Do you wish to continue? " -f Yellow -no ; Write-Host "( y = yes | n = no )" -f Gray ; brk ; brk

#Opt Out
Do {
    $Question = Read-Host "Continue with cloning of permissions?"
    If($options -notcontains $Question){Write-Host "Enter 'y' or 'n'" -f Yellow -b Black ; $chk = $false}
    Else{Switch($Question){
        'n'{brk ; brk ; Write-Host "Cancelling Script..." -f Yellow -b Black ; brk ; brk ; Exit}
        'y'{brk ; brk ; brk ; # Begin Clone Process
            # Gathering Security Groups of Cloner and Clonee
            $PermCloner = @(Get-ADUser $Cloner -Properties *).MemberOf ; $gcnt = $PermCloner.Count
            $PermClonee = @(Get-ADUser $Clonee -Properties *).MemberOf ; $cnt = 1

            # Notify Begining of Cloning
            Write-Host "Beginning Cloning Process" -f Yellow -b Black ; brk

            # Go through each Security group to be Cloned
            ForEach($p in $PermCloner){
                # Get the Group Name
                $gname = Get-ADGroup -Identity $p | Select -expand Name

                # Write Output of current group
                Write-Host "$cnt " -f Cyan -no ; Write-Host "of " -f Gray -no ; Write-Host "$gcnt " -f Cyan -no ; Write-Host ": " -f Yellow -no 
                Write-Host "$gname " -f Yellow -no ; Write-Host ": " -f Gray -no 

                # Verify if Clonee already has this security group or not
                If($Permclonee -match $gname){Write-Host "Membership already Exists!" -f Gray}
                Else{
                    # Try to add member to group. Failure means Tier 3 Permissions needed
                    Try{Add-ADGroupMember -Identity $p -Members $Clonee ; Write-Host "Member Added to Group!" -f Green}
                    Catch{Write-Host "Tier 3 Required!" -f Red ; $Tier3.Add($gname)}
                    }

                    $cnt++
                } # End of For Each Statement

            Write-Host "______________________________________________________" -f Gray ; brk ;

            # Gather Failure count and if there are any, output the amount and name of groups.
            $failed = $Tier3.Count
            If($failed -lt 1){}Else{
                Write-Host "The Following " -f Yellow -b Black -no ; Write-Host "$failed Security Groups " -f Red -b Black -no ; Write-Host "Require Tier 3 to be added to the account." -f Yellow -b Black
                ForEach($fp in $Tier3){Write-Host $fp -f DarkYellow}
                }

            $Today = Get-Date -f "dd MMM yyyy"
            $Admin = $env:USERNAME
            $Report = [pscustomobject] @{Admin = $Admin ; ClonedFrom = $Cloner ; ClonedTo = $Clonee ; Date = $Today}
            $Report | Select Admin,ClonedFrom,ClonedTo,Date | Export-Csv $Folder\LogFile.csv -NoTypeInformation -Encoding UTF8 -Append

            } # Endo of Cloning Process
        } # End of Switch
    } # End of 1st IF Statement
}Until($options -contains $Question)

}

Function Get-StaleUsers {
<#
.SYNOPSIS
    Grabs a list of active user accounts who have not logged on in 104+ days.
#>

<#
#--------------------------------------------------------------------------------------------------------
# Created 25 Aug 2021 by Dan Rosborough
#
# Updated on 2 April 2022 by Michael Sprous
# Update : Complete overhaul of script output and reporting system
#
# Updated on 30 Jul 2022 by Michael Sprous
# Update : Re-write of Script to include Powershell Module and add new folder structure.
#
# Updated on 29 May 2026 by Michael Sprous
# Update: Re-write of script to now gather the results quickly and display in a table format, rather than consistent output lines. More linear, less pretty.
#         Updated some of the Synopsis 
#
# This script is used to create a list of users that have not logged in for 365 days.
# These accounts should be disabled due to inactivity, IAW ISSM Policy.  The script produces 2 output files.
#
# 
# The "Users.txt" file is a list of UserIDs that will feed the second script: DisableStaleUserAccounts.
#
# Prior to running the script, update the $folder path to the specific file location.
#--------------------------------------------------------------------------------------------------------
#>


# Initial Variables
[int]$CutoffDays = 104
[string]$LogPath = "\\newton\admin\Scripts\Monthly Scripts\Stale User Accounts\Logs"

try {
    $CutoffDate = (Get-Date).AddDays(-$CutoffDays)
    $Days = 90
    
    # Setup Paths
    $TimeStamp = Get-Date -Format "yyyyMMddHHmmss"
    $LogFolder = Join-Path $LogPath "$Days Days - $(Get-Date -Format 'dd MMM yyyy')"
    $null = New-Item -ItemType Directory -Path $LogFolder -Force
    $UserTxtPath = Join-Path $LogFolder "Users.txt"
    $CsvPath     = Join-Path $LogFolder "StaleUsers-$TimeStamp.csv"

    # Define all exclusions in a single, clean Regex string for fast filtering
    $Exclusions = '(?i)^(gold|svc-|ALPINE-|-NSSE|TCC|SIMCELL|DECCSSP|Exch|maxreg|mxintadm|helpdesk|newton |.*-ad|.*admin.*|.*-dto)'

    Write-Host "Querying AD for users inactive since $CutoffDate..." -ForegroundColor Cyan

    # 1. Get Enabled users where Password is NOT set to never expire
    # 2. Filter locally for Dates and Exclusions
    $StaleUsers = Get-ADUser -Filter {Enabled -eq $true -and PasswordNeverExpires -eq $false} -Properties LastLogonDate, whenCreated, Description | 
        Where-Object {
            $EffectiveDate = if ($_.LastLogonDate) { $_.LastLogonDate } else { $_.whenCreated }
            ($EffectiveDate -lt $CutoffDate) -and ($_.SamAccountName -notmatch $Exclusions)
        }

    if (-not $StaleUsers) {
        Write-Warning "No stale users found."
        return
    }

    # Generate Output Data
    $Report = $StaleUsers | Select-Object Name, SamAccountName, Description,
        @{N='LastLogon'; E={ 
            if ($_.LastLogonDate) { $_.LastLogonDate } else { $_.whenCreated } 
        }},
        @{N='DaysInactive'; E={
            $TargetDate = if ($_.LastLogonDate) { $_.LastLogonDate } else { $_.whenCreated }
            ((Get-Date) - $TargetDate).Days
        }} | 
    Sort-Object DaysInactive -Descending


    # Console Output (Simple/Clean)
    $Report | Format-Table Name, SamAccountName, LastLogon, DaysInactive -AutoSize

    # File Exports
    $Report.SamAccountName | Out-File -FilePath $UserTxtPath -Encoding UTF8
    $Report | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

    Write-Host ("`nFound {0} users." -f $Report.Count) -ForegroundColor Green
    Write-Host ("Saved to: {0} `n" -f $UserTxtPath) 
} catch {
    Write-Error "Error retrieving users: $_"
}

}

Function Get-StaleAdmins {
<#
 .Synopsis
  Identifies and reports administrator accounts that have been inactive for more than 30 days, supporting security compliance with ISSM policies.

 .Description
  This script identifies administrator accounts (containing "-admin" in their username) that have not been logged into for at least 30 days. Due to Active Directory replication delays, the actual cutoff date used is 44 days to ensure accuracy.
  
  The script generates a CSV report with detailed account information and a text file containing the usernames of stale accounts, which can be used as input for the companion script 'Disable-StaleAdmins'.
  
  Results are stored in a dated directory for record-keeping and compliance documentation.

 .Parameter OpenReport
  When specified, automatically opens the generated CSV report and text file after script execution.

 .Example
    Get-StaleAdmins
    # Runs the script with the default cutoff date of 44 days (accounting for replication) and outputs results to the console and log files.
 
 .Example
    Get-StaleAdmins -OpenReport
    # Runs the script and automatically opens the generated CSV report and text file after completion.
   
 .Notes
    Author: Michael Sprous
    Created: 08 Oct 2022
    Purpose: Security compliance with ISSM policy requiring inactive admin accounts to be identified and disabled.
#>


#--------------------------------------------------------------------------------------------------------
# Created 08 Oct 2022 by Michael Sprous
#
# Get-StaleAdmins script is used to create a list of users that have not logged in for 30 days.
# These accounts should be disabled due to inactivity, IAW ISSM Policy.
#
# This script is used to determin the status of Admins Accounts in Active Directory.
# Per the ISSM, Account not logged into in the las 90 days will be diabled.
# In Active Directory, Last LogonDate can be up to 14 days off due to replication.
# Therefore, the cutoff for disableing accounts is 104 days.
#
# It should be moved to the High side and sent to the ISSM and his alternate when the work is completed.
# 
# The "Users.txt" file is a list of UserIDs that will feed the second script: Disable-StaleAdmins.
#
# Updated on 29 May 2026 by Michael Sprous
# Update: Re-write of script to now gather the results quickly and display in a table format, rather than consistent output lines. More linear, less pretty.
#         Updated some of the Synopsis.
#--------------------------------------------------------------------------------------------------------

# Initial Variables
[int]$CutoffDays = 104
[string]$LogPath = "\\newton\admin\scripts\Monthly Scripts\Stale Admin Accounts\Logs"

try {
    $cutOffDate = (Get-Date).AddDays(-$cutOffDays)
    $Days = 90

    # Setup Paths
    $TimeStamp = Get-Date -Format "yyyyMMddHHmmss"
    $LogFolder = Join-Path $LogPath "$Days Days - $(Get-Date -Format 'dd MMM yyyy')"
    $null = New-Item -ItemType Directory -Path $LogFolder -Force
    $UserTxtPath = Join-Path $LogFolder "Admins.txt"
    $CsvPath = Join-Path $LogFolder "StaleAdmins-$TimeStamp.csv"

    # Define all exclusions in a single regex string for fast filtering
    $Exclusions = '(?i)^(.*-ad|.*-admin.*|.*-dto)'

    Write-Host "Querying AD for admin accounts inactive since $cutOffDate..." -ForegroundColor Cyan

    # 1. Get enabled admins where password is NOT set to never expire
    # 2. Filter locally for dates and exclusions
    $StaleAdmins = Get-ADUser -Filter {Enabled -eq $true -and PasswordNeverExpires -eq $false} -Properties LastLogonDate, whenCreated, Description |
        Where-Object {
            $EffectiveDate = if ($_.LastLogonDate) { $_.LastLogonDate } else { $_.whenCreated }
            ($EffectiveDate -lt $cutOffDate) -and ($_.SamAccountName -match $Exclusions)
        }
    
    if (-not $StaleAdmins) {
        Write-Warning "No stale admin accounts found."
        return
    }

    $Report = $StaleAdmins | Select-Object Name, SamAccountName, Description,
        @{N='LastLogon'; E={ if ($_.LastLogonDate) { $_.LastLogonDate } else { $_.whenCreated }}},
        @{N='DaysInactive'; E={ $TargetDate = if ($_.LastLogonDate) { $_.LastLogonDate } else { $_.whenCreated } ((Get-Date) - $TargetDate).Days }} |
    Sort-Object DaysInactive -Descending

    # Console output (Simple/clean)
    $Report | Format-Table Name, SamAccountName, LastLogon, DaysInactive -AutoSize

    # File Exports
    $Report.SamAccountName | Out-File -FilePath $UserTxtPath -Encoding utf8
    $Report | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

    # Final output
    Write-Host ("`nFound {0} stale admin acounts." -f $Report.count) -ForegroundColor Green
    Write-Host ("Saved to: {0} `n" -f $UserTxtPath)

}
catch {
    Write-Error "Error retrieving accounts: $_"
}

# Logging of script
$logpath = "\\newton\admin\SysAdmin Powershell Module\Scripts\BattleRythm\Logs\Log_BattleRhythm.csv"
Add-Content -Path $logpath -Value "`Get-StaleAdmins,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami)"

}

Function Get-DomainPrinters {


<#
 .Synopsis
  Retrieve all Printer information for every Print Server on the domain 

 .Description
  Allows the Administrator to easily access Printer information of all Print Servers connected on the domain

 .Example
    Get-DomainPrinters
    # Runs the Script normally.

 .Example
    Get-DomainPrinters 
    # Opens the created report file after completion.
   
#>

$fpServers = (Get-ADComputer -filter "operatingsystem -like '*server*' -and name -like '*FP03' -or name -like '*FP07'").Name 
$folder = "\\newton.pentagon.mil\admin\scripts\Monthly Scripts\Printers\Reports"
$LogDate = Get-Date -Format "dd MMM yyyy"
$LogTime = Get-Date -Format "hhmmss"
$Logfile = "$folder\Printers-$LogDate-$LogTime.csv"

[int]$scnt = 0
[int]$pcnt = 0
[int]$qsOnline = 0
[int]$qsBadPort = 0
[int]$qsShared = 0
[int]$qsPublished = 0

[int]$TotalSvrs = $fpServers.count 
$AllPrinters = @()
$Report = @()

Write-Host ""
Write-Host " ==== BEGINNING TO GRAB PRINTERS ==== " -f Yellow -b Black
Write-Host ""

ForEach($svr in $fpServers){
    $scnt++
    Write-Host "$scnt/$TotalSvrs - " -f Gray -no 
    Write-Host "Gathering Printers from $svr..." -no 
    If(!(Test-Connection -ComputerName $svr -Quiet -count 2 -BufferSize 16)){ Write-Host "Offline." -f Red ; Continue }
    
    Try { 
        $Printers = Invoke-Command -ComputerName $svr -ScriptBlock { @(Get-Printer | where PortName -ne 'PortPrompt:') } -ErrorAction Si 
        $AllPrinters += $Printers
    } Catch { Write-Warning "Could not Gather Printers. Check Spooler." ; Continue }

    Write-Host "Done." -f Green
}

Write-Host ""
Write-Host " ==== BEGINNING TO GO THROUGH EACH PRINTER ==== " -f Yellow -b Black
Write-Host ""

ForEach($printer in $AllPrinters){
    $pcnt++ 

    $obj = New-Object PSObject 
    $obj | Add-Member NoteProperty -Name "ServerName" -Value "$($Printer.PSComputerName)"
    $obj | Add-Member NoteProperty -Name "PrinterName" -Value "$($Printer.Name)"
    $obj | Add-Member NoteProperty -Name "Status" -Value $null
    $obj | Add-Member NoteProperty -Name "Port" -Value "$($Printer.PortName)"
    $obj | Add-Member NoteProperty -Name "Shared" -Value "$($Printer.Shared)"
    $obj | Add-Member NoteProperty -Name "ShareName" -Value "$($Printer.ShareName)"
    $obj | Add-Member NoteProperty -Name "Published" -Value "$($Printer.Published)"
    $Obj | Add-Member NoteProperty -Name "DriverName" -Value "$($Printer.DriverName)"


    Write-Host "$pcnt/$($AllPrinters.count) - " -f Gray -no  ; Write-Host "$($obj.PrinterName) " -no ; Write-Host "on " -f Gray -no ; Write-Host "$($obj.ServerName)" -no ; Write-Host " ... " -f Gray -no 
    Write-Host "$($obj.Port)" -no  ; Write-Host " - " -f Gray -no
    If(!(($ping = ping $($printer.PortName) -n 3) -match "Reply from")){ $obj.Status = "Offline" ; Write-Host "Offline" -f Red } Else{ $obj.Status = "Online" ; Write-Host "Online" -f Green } 
    

    If($obj.Status -eq "Online"){ $qsOnline++ }
    If($obj.Shared -eq "True"){ $qsShared++ }
    If($obj.Published -eq "True"){ $qsPublished++ }
    If(!($obj.Port -as [ipaddress] -as [bool])){ $qsBadPort++ }

    $Report += $obj 
}

Write-Host "" 
Write-Host " ==== QUICK STATS ==== " -f Yellow -b Black
Write-Host ""
Write-Host "Printer Count: " -no ; Write-Host "$($AllPrinters.count) Printers" -no  ; Write-Host " on " -f Gray -no ; Write-Host "$($fpServers.count) Servers"
Write-Host "Printers Online: " -no ; Write-Host "$qsOnline" -no ; Write-Host " of "  -f Gray -no ; Write-Host "$($AllPrinters.count)"
Write-Host "Printers Shared: " -no ; Write-Host "$qsShared" -no ; Write-Host " of " -f Gray -no ; Write-Host "$($AllPrinters.count)"
Write-Host "Printers Published: " -no ; Write-Host "$qsPublished" -no ; Write-Host " of " -f Gray -no ; Write-Host "$($AllPrinters.count)"
Write-Host "Printers with BAD Ports : " -no ; Write-Host "$qsBadPort" -no ; Write-Host " of " -f Gray -no ; Write-Host "$($AllPrinters.Count)"

Write-Host ""
$Report | Export-Csv -Path $Logfile -NoTypeInformation -Encoding UTF8
Write-Host "Report saved to $Logfile" 

Write-Host "Opening report...." 
Invoke-Item $Logfile
Write-Host "Script Complete." -f Green


$logpath = "\\newton\admin\SysAdmin Powershell Module\Scripts\BattleRythm\Logs\Log_BattleRhythm.csv"
Add-Content -Path $logpath -Value "`Disable-StaleUsers,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami)"
}

Function Reset-BitLockerKey {
<#
 .Synopsis
  Automates any type of Datacall for users and computers 

 .Description
  Allows admins to quickly reset BitLocker Encryption on any machine(s)

 .Example
   Reset-BitlockerKey
   # Load the default script
   

#>

# The commands that are in use are listed below. These commands complete the process of Disabling and Enabling the encryption on a drive. 
# After this completes, it will reveal the new recovery password
# - Disable-BitLocker : This disables the encryption on the drive
# - Enable-BitLocker  : This enables the encryption and resets the recovery key.

# Variables to Load at beginning
[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
$Options = @('1','2','I','i','Q','q')

Function Load-Menu {

    Write-Host "" # Creating a gap within the console

    Write-Host " [Main Options] " -f Cyan
    Write-Host "1. " -f Gray -no ; Write-Host "Reset Recovery Key" -f Yellow ; 
    Write-Host "2. " -f Gray -no ; Write-Host "Retreive Recovery Key" -f Yellow
    Write-Host "" # Creating a gap within the console

    Write-Host " [Other Options]" -f Cyan
    Write-Host "I. " -f Gray -no ; Write-Host "Instructions" -f Yellow
    Write-Host "Q. " -f Gray -no ; Write-Host "Exit" -f Yellow
       
}

Function Menu-Choice {
    Param([Switch]$Reset,[Switch]$Retrieve,[switch]$fpnfo)

    If($Reset){
        $Comp = Read-Host "Please Enter the Hostname of the Computer" ; Clear-Host
        Write-Host "==== " -f Gray -no ; Write-Host $Comp -f Cyan -no ; Write-Host " ====" -f Gray

        # This is the connection test. If the connection fails, it will go back to the main menu
        Write-Host "Connection Test : " -f Gray -no
        If(!(Test-Connection $Comp -Count 2 -BufferSize 16 -ErrorAction Si)){Write-Host "Offline." -f Red ; Start-Sleep -sec 2 ; Load-Menu}
        Else{
             Write-Host "Online." -f Green

             # This is the WinRM Service Check. WinRM Needs to be running in order to establish a remote connection
             Write-Host "Remote Connection : " -f Gray -no ; 
             If(!(Get-Service WinRM -ComputerName $Comp -ErrorAction Si)){Write-Host "Cannot Establish." -f Red ; Start-Sleep -sec 2 ; Load-Menu}
             Else{
                  $winrm = Get-Service WinRM -ComputerName $Comp -ErrorAction Si | Select Status
                  If($winrm -eq "Running"){}Else{$Start = Get-Service WinRM -ComputerName $Comp | Start-Service}
                  Write-Host "Established." -f Green

                  # This begins the decryption of BitLocker
                  Write-Host "Decrypting Drive.." -f Gray -No
                  Invoke-Command -ComputerName $Comp {Disable-BitLocker -MountPoint C} -AsJob
                  Do{
                     $Status = Invoke-Command -ComputerName $Comp {(Get-BitLockerVolume -MountPoint C).VolumeStatus}
                     If($Status -eq "DecryptionInProgress"){Write-Host "." -f Gray -no}
                     ElseIf($Status -eq "FullyDecrypted"){Write-Host "Done." -f Green} ; Start-Sleep -sec 5
                     }
                  Until($Status -eq "FullyDecrypted")

                  # This begins Encryption of BitLocker
                  Write-Host "Re-Encrypting Drive.." -f Gray -no 
                  Invoke-Command -ComputerName $Comp {Enable-BitLocker -MountPoint "C:" -EncryptionMethod XtsAes256 -RecoveryPasswordProtector -SkipHardwareTest} -AsJob
                  Do{
                     $Status = Invoke-Command -ComputerName $Comp {(Get-BitLockerVolume -MountPoint C).VolumeStatus}
                     If($Status -eq "EncryptionInProgress"){Write-Host "." -f Gray -no}
                     ElseIf($Status -eq "FullyEncrypted"){Write-Host "Done." -f Green} ; Start-Sleep -sec 5
                     }
                  Until($Status -eq "FullyEncrypted")

                  # This Gathers the correct BitLocker Key Protector
                  $BVol = Invoke-Command -ComputerName $Comp {(Get-BitLockerVolume -MountPoint C).KeyProtector | Where KeyProtectorType -Like "RecoveryPassword"}

                  # This udpates the BitLocker Key Protector to Active Directory
                  Write-Host "Updating Active Directory : " -f Gray -no 
                  Invoke-Command -ComputerName $Comp {
                    $bvol = (Get-BitLockerVolume -MountPoint C).KeyProtector | Where KeyProtectorType -Like "RecoveryPassword"
                    Backup-BitLockerKeyProtector -KeyProtectorId $bvol.KeyProtectorId
                    }
                  Write-Host "Done." -f Green

                  # This gathers the BitLocker Key Protector ID
                  Write-Host "BitLocker Key Protector ID : " -f Gray -no 
                  $KPID = $BVol.KeyProtectorId
                  Write-Host $KPID -f Yellow

                  # This Gathers the BitLocker Recovery Password
                  Write-Host "BitLocker Recovery Password : " -f Gray -No
                  $KPass = $BVol.RecoveryPassword
                  Write-Host $KPass -f Yellow
                  Write-Host ""

                  Load-Menu
                  }
            }
            
    }

    If($Retrieve){
        $Comp = Read-Host "Please Enter the Hostname of the Computer" ; Clear-Host
        Write-Host "==== " -f Gray -no ; Write-Host $Comp -f Cyan -no ; Write-Host " ====" -f Gray

        <#
        # This is the connection test. If the connection fails, it will go back to the main menu
        Write-Host "Connection Test : " -f Gray -no
        If(!(Test-Connection $Comp -Count 2 -BufferSize 16 -ErrorAction Si)){Write-Host "Offline." -f Red ; Start-Sleep -sec 2 ; Load-Menu}
        Else{
             Write-Host "Online." -f Green

             # This is the WinRM Service Check. WinRM Needs to be running in order to establish a remote connection
             Write-Host "Remote Connection : " -f Gray -no ; 
             If(!(Get-Service WinRM -ComputerName $Comp -ErrorAction Si)){Write-Host "Cannot Establish." -f Red ; Start-Sleep -sec 2 ; Load-Menu}
             Else{
                  $winrm = Get-Service WinRM -ComputerName $Comp -ErrorAction Si | Select Status
                  If($winrm -eq "Running"){}Else{$Start = Get-Service WinRM -ComputerName $Comp | Start-Service}
                  Write-Host "Established." -f Green

                  # This Gathers the correct BitLocker Key Protector
                  $BVol = Invoke-Command -ComputerName $Comp {(Get-BitLockerVolume -MountPoint C).KeyProtector | Where KeyProtectorType -Like "RecoveryPassword"}

                  # This gathers the BitLocker Key Protector ID
                  Write-Host "BitLocker Key Protector ID : " -f Gray -no 
                  $KPID = $BVol.KeyProtectorId
                  Write-Host $KPID -f Yellow

                  # This Gathers the BitLocker Recovery Password
                  Write-Host "BitLocker Recovery Password : " -f Gray -No
                  $KPass = $BVol.RecoveryPassword
                  Write-Host $KPass -f Yellow
                  Write-Host ""

                  Load-Menu
                  }
            }
            #>
        Write-Host "Retrieving Information from AD..." -f Gray -no 
        $Computer = Get-ADComputer $Comp 
        $Recovery = Get-ADObject -Filter 'ObjectClass -eq "msFVE-RecoveryInformation"' -SearchBase $Computer.DistinguishedName -Properties msFVE-RecoveryPassword | Select -expand msFVE-RecoveryPassword
        Write-Host "Done." -f Green ; Write-Host ""

        Write-Host "Recovery Password : " -f Gray -no 
        Write-Host $Recovery -f Yellow
        Write-Host ""

        Load-Menu
    }

    If($fpnfo){
        Clear-Host
        Write-Host "========== " -f Gray -no ; Write-Host "Script Instructions" -f Cyan -no ; Write-Host " ==========" -f Gray
        Write-Host ""

        # Main Instructions
        Write-Host "The Primary function of this script is to automate the BitLocker Key Reset Task." -f Gray
        Write-Host "This script has two options in which you can either Reset a BitLocker Key, or you can Retrieve a BitLocker Key." -f Gray
        Write-Host "Please be patient while the script does its job in the background." -f Gray
        Write-Host "Recommendation : " -f Green -no ; Write-Host "Retrieve the current key information first to have it as a backup incase an error occurs." -f Gray
        Write-Host "______________________________________________________________" -f Gray
        Write-Host ""

        # Resetting BitLocker Key
        Write-Host "Key Reset Instructions" -f Yellow
        Write-Host "- You will need to enter the hostname of the computer." -f Gray
        Write-Host "- If the machine is offline or cannot be remoted into, you will be sent back to the main menu." -f Gray
        Write-Host "- The script completely automates the Enabling and Disabling of BitLocker once the hostname is entered." -f Gray
        Write-Host "- Once Re-Enabled, the script will then Backup the Recovery ID and Password to Active Directory" -f Gray
        Write-Host "- Do " -f Gray -no ; Write-Host "NOT " -f Red -no ; Write-Host "cancel this script once it has started. You will break BitLocker.." -f Gray
        Write-Host ""

        # Retrieve Recovery Key
        Write-Host "Recovery Key Retrieval Instructions" -f Yellow
        Write-Host "- You will need to enter the hostname of the computer." -f Gray
        Write-Host "- If the machine is offline or cannot be remoted into, you will be sent back to the main menu." -f Gray
        Write-Host "- Once the hostname is entered, the script will provide you with the Key Protector ID and the Recovery Password." -f Gray
        Write-Host ""

        Load-Menu
    }
}


Write-Host "Welcome to the BitLocker Key Script." -f Cyan ; Write-Host "Please choose one of the options below to continue..." -f Green
Write-Host "" ; Write-Host "DO NOT CLOSE OR CANCEL UNTIL SCRIPT IS COMPLETED. YOU WILL BREAK BITLOCKER." -f Red ; Write-Host ""
Load-Menu

Do {
    $Choice = Read-Host "Select an Option"
    If($Options -notcontains $Choice){Write-Host "Invalid Option Selected.." -f Red}
    Else{
        Switch($Choice){
        '1' {Menu-Choice -Reset}
        '2' {Menu-Choice -Retrieve}
        'i' {Menu-Choice -Info}
        'q' {Exit}
            }
        }
    }Until($Choice -eq 'q')
    
    }

Function Run-HashRefresh {
<#
 .Synopsis
  Refresh the password hash for all AD Accounts requiring smartcard logon.

 .Description
  Refresh the password hash for all AD Accounts requiring smartcard logon.

 .Example
    Run-HashRefresh
    # Runs the Script as intended.
#>

$Users = Get-ADUser -Filter {(SmartcardLogonRequired -eq $True)} -Properties SmartcardLogonRequired
[int]$total = $Users.Count * 2
[int]$cnt = 0

ForEach($fp in $Users){
    $Name = $fp.Name
    $sam = $fp.SamAccountName
    
    $off = Set-ADUser -Identity $sam -SmartcardLogonRequired $false -ErrorAction si 
    $on = Set-ADUser -Identity $sam -SmartcardLogonRequired $true -ErrorAction si 

    $cnt++
    [int]$Percent = $cnt / $total * 100
    Write-Progress -Activity "Refreshing Account Hash" -Status "$Percent%" -CurrentOperation $Name -PercentComplete $Percent
}

Write-Progress -Activity "Refreshing Account Hash" -Status "$Percent%" -CurrentOperation "WAITING FOR REPLICATION TO UNLOCK ACCOUNTS" -PercentComplete $Percent
Start-Sleep -Seconds 300

ForEach($x in $Users){
    $Name = $x.Name
    $acc = $x.SamAccountName

    $unlock = Unlock-ADAccount $acc
    $cnt++
    [int]$Percent = $cnt / $total * 100
    Write-Progress -Activity "Refreshing Account Hash" -Status "$Percent%" -CurrentOperation "Unlocking $Name" -PercentComplete $Percent
}

$logpath = "\\newton\admin\SysAdmin Powershell Module\Scripts\BattleRythm\Logs\Log_BattleRhythm.csv"
Add-Content -Path $logpath -Value "`Run-HashRefresh,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami)"
}

Function Get-ExpiredCerts {


[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [switch]$AllServers
    )

# Variable configuration
$Today = Get-Date
$Days = -90
$cnt = 1
$ReportDate = Get-Date -f "dd MMM yyyy"
$Folder = '\\newton\admin\Scripts\Monthly Scripts\ExpiredCertificates'
$ReportArray = New-Object System.Collections.Generic.List[string]

# Test for -AD Credentials
[bool]$CredTested = $false
While ($CredTested -eq $false) {
    Write-Host "Both -AD and -Admin credentials are required to run this script. `nPlease enter your -AD credentials." -f Yellow -b Black
    $cred = Get-Credential -Message "Enter your -AD account credentials."

    if ($cred -eq $null) { throw "FAILED TO PROVIDE -AD CREDENTIALS" }

    Write-Host "Testing -AD Credentials. Test will NOT cause -AD account to lock out, unless you fail subsequent re-rests." -f Yellow

    Add-Type -AssemblyName System.DirectoryServices.AccountManagement
    $TestCred = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain)
    $CredTested = $TestCred.ValidateCredentials($cred.GetNetworkCredential().UserName, $Cred.GetNetworkCredential().Password)

    if ($CredTested -eq $false) {
      
        if ((Get-ADUser -Identity ($cred.GetNetworkCredential().UserName) -Properties LockedOut).LockedOut) {
            Write-Host "Your AD Account is Locked Out" 
            break
        }
      
        Write-Host "Your -AD Credentials Failed Verification.. You will be prompted to re-enter them." -f Red
    }

    if ($CredTested) { Write-Host "AD credentials PASSED verification." -f Green }
}

if ($AllServers) {
    $Servers = Get-ADComputer -Filter { (OperatingSystem -like "*WindowsServer*") -and (Name -notlike "Pizza") } | Select -expand Name 
}
else {
    $Servers = Get-ADDomainController -Filter { (OperatingSystem -like "*Windows Server*") -and (Name -notlike "Pizza") } | Select -expand Name 
}

cls

$Total = $Servers.count 

foreach ($Svr in $Servers) {
    $hasDoDcert = $false
    $certcnt = 0
    
    Write-Host "$cnt/$total - $svr"

    # Connection Test
    if (!(Test-Connection $Svr -count 2 -BufferSize 16 -Quiet -ErrorAction Si)) { Write-Host "$Svr IS OFFLINE" -f Gray }
    else { 
        

        # Connect to Servers Certificate Store
        $GatherCerts = Invoke-Command -ComputerName $Svr -Credential $cred {
            
            $Store = New-Object System.Security.Cryptography.X509Certificates.X509Store("my","LocalMachine")
            $Store.Open("ReadOnly")
            $certArray = New-Object System.Collections.Generic.List[string] 

            foreach($cert in $Store.Certificates) {
                if ($cert -ne $null) {
                    if ($cert.Issuer -like "*DOD*") {
                        
                        $hasDoDcert = $true
                        $certDetails = [pscustomobject] @{ Server="$using:svr" ; Subject="$($cert.Subject)" ; Issuer="$($cert.Issuer)" ; ExpirationDate="$([datetime]$cert.NotAfter)" ; Status="" }
                        $DaysToExpiration = ([datetime]$cert.NotAfter).AddDays($Days) 

                        if ($DaysToExpiration -gt $Today) { 
                            Write-Host "$($certDetails.Issuer) | $($certDetails.ExpirationDate) - Valid Certificate" -f Green | ft 
                            $certDetails.Status = "Valid Certificate"
                        }
                        else {
                            if ([datetime]$cert.NotAfter -lt $Today) { 
                                Write-Host "$($certDetails.Issuer) | $($certDetails.ExpirationDate) - Certificate Expired - Order new DoD Certificate" -f Red 
                                $certDetails.Status = "Certificate Expired"
                            }
                            else { 
                                Write-Host "$($certDetails.Issuer) | $($certDetails.ExpirationDate) - Certificate Expires Soon - Order a new DoD Certificate" -f Yellow 
                                $certDetails.Status = "Certificate Expires Soon"
                            }
                        }

                        $certArray += $certDetails
                            
                    }
                    
                }

            }
                
            if ($hasDoDcert -eq $false) {
                $certDetails = [pscustomobject] @{ Server="$using:svr" ; Subject="" ; Issuer="" ; ExpirationDate="" ; Status="NO DOD CERTIFICATE FOUND" }     
                Write-Host "Certficiate is missing : ********* Order New DoD Certificate *********" -f Red
                $certArray += $certDetails
            }
        
            Return $certArray
        }

        $ReportArray += $GatherCerts
        $cnt++ ;
              
    }
    Write-Host ""
}

$ReportArray | Select Server,Subject,Issuer,ExpirationDate,Status | Export-Csv -Path $Folder\$ReportDate.csv `
    -NoTypeInformation -Encoding UTF8 -Force
Invoke-Item -Path $Folder\$ReportDate.csv 
$logpath = "\\newton\admin\SysAdmin Powershell Module\Scripts\BattleRythm\Logs\Log_BattleRhythm.csv"
Add-Content -Path $logpath -Value "`Get-ExpiredCerts,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami)"

}

Function Get-CertRequest {
<#
 .Synopsis
  Create a Certificate Request for a Server. 

 .Description
  Allows the Administrator to create a Certificate request from a server and copy the CertHash to their clipboard via a GUI Form

 .Example
    Get-CertRequest -Server
    # Specifies which Server to generate the request (Mandatory)
   
#>

<#
# Created 10 Nov 2022 by Jerry Firebaugh 

# Updated on 13 Nov 2022 by Michael Sprous
# Update : Script adjusted for easier reading of code and use within powershell module

# Description ; 
This script will check if the target system is online, if so, it will check if you have admin rights on the system by attempting to connect to the \\server\C$ root share.

It will then:
Get the system GUID using a get-wmiobject command
Get the FQDN by querying DNS
Enable Windows Remote Management
Use invoke-command to run a scriptblock on the target system which generates the certificate request.
Disable Windows Remote Management
Output the Certificate Authority website URL, the certificate request, GUID and FQDN to a powershell GUI form with each value assigned a "Copy to Clipboard" button.

#>

[cmdletbinding ()]
param( [parameter (Mandatory = $true)][string]$Server)

[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
[void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')

Function brk { Write-Host "" }

<# Form Shortcut Variables #>
$Location = 'System.Drawing.Point'
$Size = 'System.Drawing.Size'
$Form = 'System.Windows.Forms.Form'
$Button = 'System.Windows.Forms.Button'
$TextBox = 'System.Windows.Forms.TextBox'
$Label = 'System.Windows.Forms.Label'
$Font1 = [System.Drawing.Font]::new("Lucida Console",8,[System.Drawing.FontStyle]::Regular)
$Font2 = [System.Drawing.Font]::new("Lucida Console",14,[System.Drawing.FontStyle]::Bold)

# Form Functions
Function Add-Label { 
Param([int]$xsize,[int]$ysize,[int]$xpos,[int]$ypos,[string]$text,[switch]$fixed3D)

$lname = New-Object $Label
$lname.Size = New-Object $Size($xsize,$ysize)
$lname.Location = New-Object $Location($xpos,$ypos)
$lname.Text = $text
If($fixed3D){$lname.BorderStyle = 'Fixed3D'}

$lname
}

Function Add-Button { 
Param([int]$xsize,[int]$ysize,[int]$xpos,[int]$ypos,[string]$text)

$bname = New-Object $Button
$bname.Size = New-Object $Size($xsize,$ysize)
$bname.Location = New-Object $Location($xpos,$ypos)
$bname.Text = $text

$bname
}

Function Add-TextBox { 
Param([int]$xsize,[int]$ysize,[int]$xpos,[int]$ypos,[string]$text)

$tbname = New-Object $TextBox
$tbname.Size = New-Object $Size($xsize,$ysize)
$tbname.Location = New-Object $Location($xpos,$ypos)
If($text -eq $null){$tbname.Text = " "} Else{$tbname.Text = $text}

$tbname
}


cls ; 

Write-Host "Cert Request Process Beginning..." -f Yellow -b Black
brk ; brk

# Update the CAsite value as new certificate authority servers come online.
# Note:  As of DEC 2025, the NPE Portal will be the required site to submit requests.
$CAsite = "https://npe-portal.csd.disa.mil/NPEPortal/#!/"

# Verify Server is Online
Write-Host "Connection Test : " -f Gray -no 
If(Test-Connection $Server -count 2 -Quiet -ErrorA Si)
{   
    Write-Host "Online." -f Green
    
    # Testing the permissions
    Write-Host "Permission Test : " -f Gray -no 
    If(!(Test-Path \\$Server\C$\ -ErrorA Si)){ Write-Host "Tier 3 Required." -f Red ; Continue }
    Else
    {
    # Gather DNS Info 
    $DNS = (Get-ADComputer $Server).DNSHostName
    Write-Host "DNS Name : " -f Gray -no ; Write-Host $DNS -f Yellow

    # Gather GUID
    $DCDN = $null
    $DCDN = @()
    $DCDN = @(Get-ADcomputer -identity "$Server")
    $DCGUID = [guid]((([directoryservices.directorysearcher] "(distinguishedname=$DCDN)").findall())[0].properties.getenumerator() | ? { $_.name -eq "objectguid"}).value[0]
    $GUID = $DCGUID.ToString("N")

    # Verify WinRM and WManSvc Services
    Write-Host "WinRM : " -f Gray -no 
    $WinRM = Set-Service WinRM -ComputerName $Server -Status Running -ErrorA Si
    If($WinRM -eq $false){ Write-Host "Failed to Start." -f Red } Else{ Write-Host "Successfully Started." -f Green }
    Write-Host "WManSvc : " -f Gray -no 
    $WManSvc = Set-Service WManSvc -ComputerName $Server -Status Running -ErrorA Si 
    If($WManSvc -eq $false){ Write-Host "Failed to Start." -f Red } Else{ Write-Host "Successfully Started." -f Green }

    # Being Invoking command
    Write-Host "Remoting to $Server to Generate Request..." -f Yellow -BackgroundColor black
    Invoke-Command -ComputerName $Server {
        $CertName = ([System.Net.Dns]::GetHostByName($env:COMPUTERNAME)).Hostname
        $FileName = $env:COMPUTERNAME
        $CSRPath = "C:\Downloads\$($FileName).txt"
        Remove-Item $CSRPath -ErrorA Si 
        $fpNFPath = "C:\Downloads\$($FileName).inf"
        $Signature = '$Windows NT$'
        
        # The INF file options were taken out of the "Domain Controller Certificate Request Generation NIPR Download" on https://cyber.mil/pki-pke/tools-configuration-files/
        # The official domain controller certificate creation powershell script.  That script is designed to be run from a local session and has many options we don't need.

        $fpNF =
@"
[Version]
Signature = "$Signature" 
[NewRequest]
Subject = "CN=$CertName,OU=DISA,OU=PKI,OU=DoD,O=U.S. Government,C=US"
KeySpec = 1
KeyLength = 2048
Exportable = TRUE
MachineKeySet = TRUE
ProviderName = "Microsoft software Key Storage provider"
ProviderType = 12
RequestType = PKCS10
"@  
        Write-Host "Certificate Request is being Generated." -f Yellow
        $fpNF | Out-File -FilePath $fpNFPath -Force
        certreq -new $fpNFPath $CSRPath
        } # End of Invoke Command

    Write-Host "Certificate Request has been Generated." -f Green ; brk
    $CertHash = (Get-Content "\\$Server\C$\Downloads\$Server.txt")

    # Stopping WinRM
    $stop = Get-Service WinRM -ComputerName $Server | Stop-Service -Force

    # Building the GUI Form
    $MainForm = New-Object $Form ; $MainForm.Size = New-Object $Size(700,625) ; $MainForm.Text = "Copy-Paste Values"
    $CloseButton = Add-Button -xsize 140 -ysize 33 -xpos 275 -ypos 540 -text "CLOSE" ; $CloseButton.Font = $Font2 
    $CloseButton.BackColor = 'Pink' ; $CloseButton.DialogResult = 'OK' ; $CloseButton.Add_Click({$MainForm.Close()})
    
    $CASiteLabel = Add-Label -xsize 200 -ysize 20 -xpos 10 -ypos 15 -text "CA Request URL:"
    $CASiteTextBox = Add-TextBox -xsize 500 -ysize 25 -xpos 10 -ypos 35 -text $CASite ; $CASiteTextBox.Font = $Font1
    $CASiteButton = Add-Button -xsize 120 -ysize 23 -xpos 525 -ypos 33 -text "Copy to Clipboard" ; $CASiteButton.Add_Click({Set-Clipboard $CASite ; $CASiteTextBox.ForeColor = 'Green'})
    
    $CertReqLabel = Add-Label -xsize 200 -ysize 20 -xpos 10 -ypos 65 -text "Certificate Request Text"
    $CertReqTextBox = Add-TextBox -xsize 500 -ysize 325 -xpos 10 -ypos 85 -text $CertHash ; $CertReqTextBox.Font = $Font1 ;  $CertReqTextBox.multiline = $true
    $CertReqButton = Add-Button -xsize 120 -ysize 23 -xpos 525 -ypos 150 -text "Copy to Clipboard" ; $CertReqButton.Add_Click({Set-Clipboard $CertHash ; $CertReqTextBox.ForeColor = 'Green'})
    
    $FQDNLabel = Add-Label -xsize 200 -ysize 20 -xpos 10 -ypos 430 -text "FQDN / DNS Name:"
    $FQDNTextBox = Add-TextBox -xsize 500 -ysize 25 -xpos 10 -ypos 450 -text $DNS ; $FQDNTextBox.Font = $Font1
    $FQDNButton = Add-Button -xsize 120 -ysize 23 -xpos 525 -ypos 448 -text "Copy to Clipboard" ; $FQDNButton.Add_Click({Set-Clipboard $DNS ; $FQDNTextBox.ForeColor = 'Green'})

    $GUIDLabel = Add-Label -xsize 200 -ysize 20 -xpos 10 -ypos 490 -text "GUID:"
    $GUIDTextBox = Add-TextBox -xsize 500 -ysize 25 -xpos 10 -ypos 510 -text $GUID ; $GUIDTextBox.Font = $Font1
    $GUIDButton = Add-Button -xsize 120 -ysize 23 -xpos 525 -ypos 508 -text "Copy to Clipboard" ; $GUIDButton.Add_Click({Set-Clipboard $GUID ; $GUIDTextBox.ForeColor = 'Green'})

    $Objects = @($CASiteLabel,$CASiteTextBox,$CASiteButton,$CertReqLabel,$CertReqTextBox,$CertReqButton,$FQDNLabel,$FQDNTextBox,$FQDNButton,$GUIDLabel,$GUIDTextBox,$GUIDButton,$CloseButton)
    $MainForm.Controls.AddRange($Objects)
    [System.Windows.Forms.Application]::Run($MainForm)
    }
            
}
Else # Server not Online
{
    Write-Host "Error." -f Red
    brk ; brk
    Write-Host "Check spelling, DNS name resolution, and/or troubleshoot network connectivity..." -f Yellow -b Black
    brk ; brk
}

}

Function Get-UptimeReport {
[cmdletbinding()]
Param ( 
[parameter (Mandatory = $false)] [switch]$AllDCs, 
[parameter (Mandatory = $false)] [switch]$MemberServers,
[parameter (Mandatory = $false)] [switch]$AllServers
)

$noswitch = $true
$cnt = 0

Function getuptime {
$Servers = $Servers | Sort-Object
$Total = $Servers.Count

    ForEach($Svr in $Servers){
        $cnt = $cnt + 1
        $Lastreboottime = $null
        Write-Host "Server " -f Gray -no ; Write-Host "$cnt" -f Yellow -no ; Write-Host " of " -f Gray -no ; Write-Host $Total -f Yellow -no ; Write-Host " - " -f Gray -no ; Write-Host $Svr -f Cyan -no
        # Connection Test
        If(!(Test-Connection $Svr -Count 2 -BufferSize 16 -Quiet -ErrorAction Si)){Write-Host " - Offline." -f Red}
        Else{
        Try{
        $Lastreboottime = Get-WmiObject -ComputerName $Svr win32_operatingsystem -ErrorAction silentlycontinue | select @{LABEL='LastBootUpTime';EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}}
        $OS = Get-WmiObject -ComputerName $Svr Win32_OperatingSystem -ErrorA Si 
        $Timespan = $OS.ConvertToDateTime($OS.LocalDateTime) - $OS.ConvertToDateTime($OS.LastBootUpTime)
        [int]$uptime = "{0:00}" -f $Timespan.TotalHours
        write-host " - Last Bootup Time: " -f Gray -no
        Write-Host $Lastreboottime.LastBootUpTime.GetDateTimeFormats('G')[0] -f Green -no 
        Write-Host " | " -f Gray -no ; Write-Host "$uptime hours" -f Green
        }
        Catch{Write-Host " - Error. Unable to query last boot time." -f Red}
        
        }
    }
}


If($MemberServers){
$noswitch = $false
$Servers = Get-ADComputer -Filter {(OperatingSystem -like "*Windows Server*") -and (Name -notlike "Pizza")} | Select -expand Name
$XDCs = $Servers | Where-Object { $_ -like "*XDC*" }
$Servers = $Servers | Where-Object {  $_ -notlike "*DC*" }
$Servers = $Servers + $XDCs
getuptime
}

If($AllServers){
$noswitch = $false
$Servers = Get-ADComputer -Filter {(OperatingSystem -like "*Windows Server*") -and (Name -notlike "Pizza")} | Select -expand Name
getuptime
}


if($AllDCs){
$noswitch = $false
$Servers = Get-ADDomainController -Filter {(OperatingSystem -like "*Windows Server*2019*") -and (Name -notlike "Pizza")} | Select -expand Name
getuptime
}

if($noswitch){
Write-host "Error.  Use -AllDCs, -MemberServers, or -AllServers switch after the get-UptimeReport command" -f red
Write-host "Example:  Get-UptimeReport -AllDCs" -f Yellow
}

$logpath = "\\newton\admin\SysAdmin Powershell Module\Scripts\BattleRythm\Logs\Log_BattleRhythm.csv"
Add-Content -Path $logpath -Value "`Get-UptimeReport,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami)"
}

Function Migrate-Printers {
[cmdletbinding()]
param (
[parameter (mandatory = $true)] [string]$From,
[parameter (mandatory = $true)] [string]$To)

#break line
Function brk {Write-Host ""}

#Verify if servers are online
Write-Host "Verifying Server connections..." -f Yellow -b Black
Write-Host "$From ... " -f Gray -no ; If(!(Test-Connection $From -Count 2 -BufferSize 16 -Quiet -ErrorAction Si)){Write-Host "Offline." -f Red ; $gtg1 = $false} Else{Write-Host "Online." -f Green ; $gtg1 = $true}
Write-Host "$To ... " -f Gray -no ; If(!(Test-Connection $To -Count 2 -BufferSize 16 -Quiet -ErrorAction Si)){Write-Host "Offline." -f Red ; $gtg2 = $false} Else{Write-Host "Online." -f Green ; $gtg2 = $true}

If($gtg1 -eq $false -or ($gtg2 -eq $false)){ brk
    Write-Host "SERVER CONNECTION FAILED. PLEASE MAKE SURE SERVERS ARE ONLINE OR TYPED CORRECTLY AND TRY AGAIN!" -f Red -b Black ; Exit
    }
Else{

brk ; Write-Host "BEGINNING MIGRATION....PLEASE BE PATIENT..." -f Yellow -b Black ; start-sleep -Seconds 1


Invoke-Command -comp $from { c:\windows\system32\spool\tools\printbrm -b -s \\$using:from -f c:\Downloads\ExportofPrinters.printerexport }

$Copy = Copy-Item -Path \\$from\c$\Downloads\ExportofPrinters.printerexport -Destination \\$to\c$\Downloads\

Invoke-Command -comp $to { c:\windows\system32\spool\tools\printbrm -r -f c:\Downloads\ExportofPrinters.printerexport }



brk ; brk
Write-Host "NOTE: IF ANY SPECIFC PRINTER FAILED, PLEASE MIGRATE THAT PRINTER MANUALLY. IT IS A KNOWN ISSUE. -MIKE SPROUS" -f Yellow -b Black
Write-Host "Printer Migration Complete." -f Green

}
}

Function Remote-EMS {
# ======================================================================
# Name: Remote Exchange Management Shell
# Author: Al Stewart and Michael Sprous
# Date: 21 May 2023
# Purpose: Create a local Exchange Management Shell on your workstation
# ======================================================================

If(!($Get = Get-PSSession -ErrorAction Si | Where ConfigurationName -like "Microsoft.Exchange")){ 
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://n490mbx01.newton.pentagon.mil/powershell | Import-PSSession $Session  }
Else{ Write-Host "Session Already Exists!" -f Yellow -b Black }
    
}

Function Find-InDHCP {

[cmdletbinding()]
Param(
    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    $System
)

$dSvr = @()
$dSvr += "N020DHCP01"
$dSvr += "N490DHCP01"

Foreach ($dhcp in $dSvr) {
Write-host ""
Write-host "Searching $dhcp ........"
Write-host ""
$allscope = Get-DhcpServerv4Scope -ComputerName $dhcp

   foreach ($ScopeID in $allscope) 
    {
    Get-DhcpServerv4Lease -ComputerName $dhcp -ScopeId $ScopeID.scopeid | where {($_.hostname -like $System) -or ($_.ClientId -match $System) -or ($_.IPAddress -eq $System)}
    }
}

}

Function Test-ExchangeServerHealth {

<#
.SYNOPSIS
Test-ExchangeServerHealth.ps1 - Exchange Server Health Check Script.

Get-Service | Where {$_.DisplayName -Like "*Exchange*"} | ft DisplayName, Name, Status

.DESCRIPTION 
Performs a series of health checks on Exchange servers and DAGs
and outputs the results to screen, and optionally to log file, HTML report,
and HTML email.

Use the ignorelist.txt file to specify any servers, DAGs, or databases you
want the script to ignore (eg test/dev servers).

.OUTPUTS
Results are output to screen, as well as optional log file, HTML report, and HTML email

.PARAMETER Server
Perform a health check of a single server

.PARAMETER ReportMode
Set to $true to generate a HTML report. A default file name is used if none is specified.

.PARAMETER ReportFile
Allows you to specify a different HTML report file name than the default.

.PARAMETER SendEmail
Sends the HTML report via email using the SMTP configuration within the script.

.PARAMETER AlertsOnly
Only sends the email report if at least one error or warning was detected.

.PARAMETER Log
Writes a log file to help with troubleshooting.

.EXAMPLE
.\Test-ExchangeServerHealth.ps1
Checks all servers in the organization and outputs the results to the shell window.

.EXAMPLE
.\Test-ExchangeServerHealth.ps1 -Server HO-EX2010-MB1
Checks the server HO-EX2010-MB1 and outputs the results to the shell window.

.EXAMPLE
.\Test-ExchangeServerHealth.ps1 -ReportMode -SendEmail
Checks all servers in the organization, outputs the results to the shell window, a HTML report, and

#>

#requires -version 2

[CmdletBinding()]
param (
	[Parameter( Mandatory=$false)]
	[string]$Server,

	[Parameter( Mandatory=$false)]
	[string]$ServerList,	
	
	[Parameter( Mandatory=$false)]
	[string]$ReportFile="exchangeserverhealth.html",

	[Parameter( Mandatory=$false)]
	[switch]$ReportMode,
	
	[Parameter( Mandatory=$false)]
	[switch]$SendEmail,

	[Parameter( Mandatory=$false)]
	[switch]$AlertsOnly,	
	
	[Parameter( Mandatory=$false)]
	[switch]$Log

	)

$mbxservers = @('N490MBX01.newton.pentagon.mil','N020MBX01.newton.pentagon.mil','N066MBX01.newton.pentagon.mil')
If(Test-Connection $mbxservers[0] -count 2 -quiet){$mbxremote = $mbxservers[0]}
ElseIf(Test-Connection $mbxservers[1] -count 2 -quiet){$mbxremote = $mbxservers[1]}
Else{$mbxremote = $mbxservers[-1]}
$mbxremote = 'N490MBX04.newton.pentagon.mil'
$connectionuri = "http://$($mbxremote)/powershell"


#...................................
# Variables
#...................................
$RemoteSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionuri
Import-PSSession $RemoteSession
$User = $env:USERNAME

$now = Get-Date											#Used for timestamps
$date = $now.ToShortDateString()						#Short date format for email message subject
[array]$exchangeservers = @()							#Array for the Exchange server or servers to check
[int]$transportqueuehigh = 100							#Change this to set transport queue high threshold. Must be higher than warning threshold.
[int]$transportqueuewarn = 80							#Change this to set transport queue warning threshold. Must be lower than high threshold.
$mapitimeout = 10										#Timeout for each MAPI connectivity test, in seconds
$pass = "Green"
$warn = "Yellow"
$fail = "Red"
$fpp = $null
[array]$serversummary = @()								#Summary of issues found during server health checks
[array]$dagsummary = @()								#Summary of issues found during DAG health checks
[array]$report = @()
[bool]$alerts = $false
[array]$dags = @()										#Array for DAG health check
[array]$dagdatabases = @()								#Array for DAG databases
[int]$replqueuewarning = 8								#Threshold to consider a replication queue unhealthy
$dagreportbody = $null

$myDir = "\\newton\admin\SysAdmin Powershell Module\Scripts\Test-ExchangeServerHealth"

#...................................
# Modify these Variables (optional)
#...................................

$reportemailsubject = "Exchange Server Health Report"
$fpgnorelistfile = "$myDir\ignorelist.txt"
$logfile = "$myDir\exchangeserverhealth.log"

#...................................
# Modify these Email Settings
#...................................

$smtpsettings = @{
	To =  "-nocsustainment@newton.pentagon.mil"
	From = "$User@newton.pentagon.mil"
	Subject = "$reportemailsubject - $now"
	SmtpServer = "relay.newton.pentagon.mil"
	}


#...................................
# Modify these language 
# localization strings.
#...................................

# The server roles must match the role names you see when you run Test-ServiceHealth.
$casrole = "Client Access Server Role"
$htrole = "Hub Transport Server Role"
$mbrole = "Mailbox Server Role"
$umrole = "Unified Messaging Server Role"

# This should match the word for "Success", or the result of a successful Test-MAPIConnectivity test
$success = "Success"

#...................................
# Logfile Strings
#...................................

$logstring0 = "====================================="
$logstring1 = " Exchange Server Health Check"

#...................................
# Initialization Strings
#...................................

$fpnitstring0 = "Initializing..."
$fpnitstring1 = "Loading the Exchange Server PowerShell snapin"
$fpnitstring2 = "The Exchange Server PowerShell snapin did not load."
$fpnitstring3 = "Setting scope to entire forest"

#...................................
# Error/Warning Strings
#...................................

$string0 = "Server is not an Exchange server. "
$string1 = "Server is not reachable. "
$string3 = "------ Checking"
$string4 = "Could not test service health. "
$string5 = "required services not running. "
$string6 = "Could not check queue. "
$string7 = "Public Folder database not mounted. "
$string8 = "Skipping Edge Transport server. "
$string9 = "Mailbox databases not mounted. "
$string10 = "MAPI tests failed. "
$string11 = "Mail flow test failed. "
$string12 = "No Exchange Server checks performed. "
$string13 = "Server not found in DNS. "
$string14 = "Sending email. "
$string15 = "Done."
$string16 = "------ Finishing"
$string17 = "Unable to retrieve uptime. "
$string18 = "Ping failed. "
$string19 = "No alerts found, and AlertsOnly switch was used. No email sent. "
$string20 = "You have specified a single server to check"
$string21 = "Couldn't find the server $server. Script will terminate."
$string22 = "The file $fpgnorelistfile could not be found. No servers, DAGs or databases will be ignored."
$string23 = "You have specified a filename containing a list of servers to check"
$string24 = "The file $serverlist could not be found. Script will terminate."
$string25 = "Retrieving server list"
$string26 = "Removing servers in ignorelist from server list"
$string27 = "Beginning the server health checks"
$string28 = "Servers, DAGs and databases to ignore:"
$string29 = "Servers to check:"
$string30 = "Checking DNS"
$string31 = "DNS check passed"
$string32 = "Checking ping"
$string33 = "Ping test passed"
$string34 = "Checking uptime"
$string35 = "Checking service health"
$string36 = "Checking Hub Transport Server"
$string37 = "Checking Mailbox Server"
$string38 = "Ignore list contains no server names."
$string39 = "Checking public folder database"
$string40 = "Public folder database status is"
$string41 = "Checking mailbox databases"
$string42 = "Mailbox database status is"
$string43 = "Offline databases: "
$string44 = "Checking MAPI connectivity"
$string45 = "MAPI connectivity status is"
$string46 = "MAPI failed to: "
$string47 = "Checking mail flow"
$string48 = "Mail flow status is"
$string49 = "No active DBs"
$string50 = "Finished checking server"
$string51 = "Skipped"
$string52 = "Using alternative test for Exchange 2013 CAS-only server"
$string60 = "Beginning the DAG health checks"
$string61 = "Could not determine server with active database copy"
$string62 = "mounted on server that is activation preference"
$string63 = "unhealthy database copy count is"
$string64 = "healthy copy/replay queue count is"
$string65 = "(of"
$string66 = ")"
#$string67 = "unhealthy content index count is"
$string68 = "DAGs to check:"
$string69 = "DAG databases to check"



#...................................
# Functions
#...................................

#This function is used to generate HTML for the DAG member health report
Function New-DAGMemberHTMLTableCell()
{
	param( $lineitem )
	
	$htmltablecell = $null

	switch ($($line."$lineitem"))
	{
		$null { $htmltablecell = "<td>n/a</td>" }
		"Passed" { $htmltablecell = "<td class=""pass"">$($line."$lineitem")</td>" }
		default { $htmltablecell = "<td class=""warn"">$($line."$lineitem")</td>" }
	}
	
	return $htmltablecell
}

#This function is used to generate HTML for the server health report
Function New-ServerHealthHTMLTableCell()
{
	param( $lineitem )
	
	$htmltablecell = $null
	
	switch ($($reportline."$lineitem"))
	{
		$success {$htmltablecell = "<td class=""pass"">$($reportline."$lineitem")</td>"}
        "Success" {$htmltablecell = "<td class=""pass"">$($reportline."$lineitem")</td>"}
        "Pass" {$htmltablecell = "<td class=""pass"">$($reportline."$lineitem")</td>"}
		"Warn" {$htmltablecell = "<td class=""warn"">$($reportline."$lineitem")</td>"}
		"Access Denied" {$htmltablecell = "<td class=""warn"">$($reportline."$lineitem")</td>"}
		"Fail" {$htmltablecell = "<td class=""fail"">$($reportline."$lineitem")</td>"}
        "Could not test service health. " {$htmltablecell = "<td class=""warn"">$($reportline."$lineitem")</td>"}
		"Unknown" {$htmltablecell = "<td class=""warn"">$($reportline."$lineitem")</td>"}
		default {$htmltablecell = "<td>$($reportline."$lineitem")</td>"}
	}
	
	return $htmltablecell
}

#This function is used to write the log file if -Log is used
Function Write-Logfile()
{
	param( $logentry )
	$timestamp = Get-Date -DisplayHint Time
	"$timestamp $logentry" | Out-File $logfile -Append
}

#This function is used to test service health for Exchange 2013 CAS-only servers
Function Test-E15CASServiceHealth()
{
	param ( $e15cas )
	
	$e15casservicehealth = $null
	$servicesrunning = @()
	$servicesnotrunning = @()
	$casservices = @(
		"IISAdmin",
		"W3Svc",
		"WinRM",
		"MSExchangeADTopology",
		"MSExchangeDiagnostics",
		"MSExchangeFrontEndTransport",
		#"MSExchangeHM",
		"MSExchangeIMAP4",
		"MSExchangePOP3",
		"MSExchangeServiceHost",
		"MSExchangeUMCR"
		)
		
	try {
		$servicestates = @(Get-WmiObject -ComputerName $e15cas -Class Win32_Service -ErrorAction STOP | where {$casservices -icontains $_.Name} | select name,state,startmode)
	}
	catch
	{
		if ($Log) {Write-LogFile $_.Exception.Message}
		Write-Warning $_.Exception.Message
		$e15casservicehealth = "Fail"
	}	
	
	if (!($e15casservicehealth))
	{
		$servicesrunning = @($servicestates | Where {$_.StartMode -eq "Auto" -and $_.State -eq "Running"})
		$servicesnotrunning = @($servicestates | Where {$_.Startmode -eq "Auto" -and $_.State -ne "Running"})
		if ($($servicesnotrunning.Count) -gt 0)
		{
			Write-Verbose "Service health check failed"
		    Write-Verbose "Services not running:"
		    foreach ($service in $servicesnotrunning)
		    {
		        Write-Verbose "- $($service.Name)"	
		    }
			$e15casservicehealth = "Fail"	
		}
		else
		{
			Write-Verbose "Service health check passed"
			$e15casservicehealth = "Pass"
		}
	}
	return $e15casservicehealth
}

#This function is used to test mail flow for Exchange 2013 Mailbox servers
Function Test-E15MailFlow()
{
	param ( $e15mailboxserver )

	$e15mailflowresult = $null
	
	Write-Verbose "Creating PSSession for $e15mailboxserver"
    $url = (Get-PowerShellVirtualDirectory -Server $e15mailboxserver -AdPropertiesOnly | Where {$_.Name -eq "Powershell (Default Web Site)"}).InternalURL.AbsoluteUri
    if ($url -eq $null)
    {
        $url = "http://$e15mailboxserver/powershell"
    }

	try
	{
	    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $url -ErrorAction STOP
	}
	catch
	{
	    Write-Verbose "Something went wrong"
		if ($Log) {Write-LogFile $_.Exception.Message}
    	Write-Warning $_.Exception.Message
		$e15mailflowresult = "Fail"
	}

	try
	{
	    Write-Verbose "Running mail flow test on $e15mailboxserver"
	    $result = Invoke-Command -Session $session {Test-Mailflow} -ErrorAction STOP
	    $e15mailflowresult = $result.TestMailflowResult
	}
	catch
	{
	    Write-Verbose "An error occurred"
		if ($Log) {Write-LogFile $_.Exception.Message}
	    Write-Warning $_.Exception.Message
	    $e15mailflowresult = "Fail"
	}

	Write-Verbose "Mail flow test: $testresult"
	Write-Verbose "Removing PSSession"
	Remove-PSSession $session.Id

	return $e15mailflowresult
}

#This function is used to test replication health for Exchange 2010 DAG members in mixed 2010/2013 organizations
Function Test-e15ReplicationHealth()
{
	param ( $e15mailboxserver )

	$e15replicationhealth = $null
	
    #Find an e15 CAS in the same site
    $ADSite = (Get-ExchangeServer $e15mailboxserver).Site
    $e15cas = (Get-ExchangeServer | where {$_.IsClientAccessServer -and $_.AdminDisplayVersion -match "Version 15" -and $_.Site -eq $ADSite} | select -first 1).FQDN

	Write-Verbose "Creating PSSession for $e15cas"
    $url = (Get-PowerShellVirtualDirectory -Server $e15cas -AdPropertiesOnly | Where {$_.Name -eq "Powershell (Default Web Site)"}).InternalURL.AbsoluteUri
    if ($url -eq $null)
    {
        $url = "http://$e15cas/powershell"
    }

    Write-Verbose "Using URL $url"

	try
	{
	    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $url -ErrorAction STOP
	}
	catch
	{
	    Write-Verbose "Something went wrong"
		if ($Log) {Write-LogFile $_.Exception.Message}
    	Write-Warning $_.Exception.Message
		#$e15replicationhealth = "Fail"
	}

	try
	{
	    Write-Verbose "Running replication health test on $e15mailboxserver"
	    #$e15replicationhealth = Invoke-Command -Session $session {Test-ReplicationHealth} -ErrorAction STOP
        $e15replicationhealth = Invoke-Command -Session $session -Args $e15mailboxserver.Name {Test-ReplicationHealth $args[0]} -ErrorAction STOP
	}
	catch
	{
	    Write-Verbose "An error occurred"
		if ($Log) {Write-LogFile $_.Exception.Message}
	    Write-Warning $_.Exception.Message
	    #$e15replicationhealth = "Fail"
	}

	#Write-Verbose "Replication health test: $e15replicationhealth"
	Write-Verbose "Removing PSSession"
	Remove-PSSession $session.Id

	return $e15replicationhealth
}


#...................................
# Initialize
#...................................

#Log file is overwritten each time the script is run to avoid
#very large log files from growing over time
if ($Log) {
	$timestamp = Get-Date -DisplayHint Time
	"$timestamp $logstring0" | Out-File $logfile
	Write-Logfile $logstring1
	Write-Logfile "  $now"
	Write-Logfile $logstring0
}

Write-Host $fpnitstring0
if ($Log) {Write-Logfile $fpnitstring0}


#Set scope to include entire forest
Write-Verbose $fpnitstring3
if ($Log) {Write-Logfile $fpnitstring3}
if (!(Get-ADServerSettings).ViewEntireForest)
{
	Set-ADServerSettings -ViewEntireForest $true -WarningAction SilentlyContinue
}


#...................................
# Script
#...................................

#Check if a single server was specified
if ($server)
{
	#Run for single specified server
	[bool]$NoDAG = $true
	Write-Verbose $string20
	if ($Log) {Write-Logfile $string20}
	try
	{
		$exchangeservers = Get-ExchangeServer $server -ErrorAction STOP
	}
	catch
	{
		#Exit because single server name was specified and couldn't be found in the organization
		Write-Verbose $string21
		if ($Log) {Write-Logfile $string21}
		Write-Error $_.Exception.Message
		EXIT
	}
}
elseif ($serverlist)
{
	#Run for a list of servers in a text file
	[bool]$NoDAG = $true
	Write-Verbose $string23
	if ($Log) {Write-Logfile $string23}
	try
	{
        $tmpservers = @(Get-Content $serverlist -ErrorAction STOP)
		$exchangeservers = @($tmpservers | Get-ExchangeServer)
    }
    catch
	{
		#Exit because file could not be found
        Write-Verbose $string24
		if ($Log) {Write-Logfile $string24}
		Write-Error $_.Exception.Message
		EXIT
    }
}
else
{
	#This is the list of servers, DAGs, and databases to never alert for
	try
	{
        $fpgnorelist = @(Get-Content $fpgnorelistfile -ErrorAction STOP)
		if ($Log) {Write-Logfile $string28}
		if ($Log) {
			if ($($fpgnorelist.count) -gt 0)
			{
				foreach ($line in $fpgnorelist)
				{
					Write-Logfile "- $line"
				}
			}
			else
			{
				Write-Logfile $string38
			}
		}
    }
    catch
	{
		Write-Warning $string22
		if ($Log) {Write-Logfile $string22}
    }
    
	#Get all servers
	Write-Verbose $string25
	if ($Log) {Write-Logfile $string25}
	$tmpservers = @(Get-ExchangeServer | sort site,name)
	
	#Remove the servers that are ignored from the list of servers to check
	Write-Verbose $string26
	if ($Log) {Write-Logfile $string26}
	foreach ($tmpserver in $tmpservers)
	{
		if (!($fpgnorelist -icontains $tmpserver.name))
		{
			$exchangeservers = $exchangeservers += $tmpserver.identity
		}
	}

	if ($Log) {Write-Logfile $string29}
	if ($Log) {
		foreach ($server in $exchangeservers)
		{
			Write-Logfile "- $server"
		}
	}
}

### Check if any Exchange 2013, 2016, or 2019 servers exist
if (Get-ExchangeServer | Where {$_.AdminDisplayVersion -like "Version 15.*"})
{
	[bool]$HasE15 = $true
}

### Begin the Exchange Server health checks
Write-Verbose $string27
if ($Log) {Write-Logfile $string27}
foreach ($server in $exchangeservers)
{
	Write-Host -ForegroundColor White "$string3 $server"
	if ($Log) {Write-Logfile "$string3 $server"}
	
	#Find out some details about the server
	try
	{
		$serverinfo = Get-ExchangeServer $server -ErrorAction Stop
	}
	catch
	{
		Write-Warning $_.Exception.Message
		if ($Log) {Write-Logfile $_.Exception.Message}
		$serverinfo = $null
	}

	if ($serverinfo -eq $null )
	{
		#Server is not an Exchange server
		Write-Host -ForegroundColor $warn $string0
		if ($Log) {Write-Logfile $string0}
	}
	elseif ( $serverinfo.IsEdgeServer )
	{
		Write-Host -ForegroundColor White $string8
		if ($Log) {Write-Logfile $string8}
	}
	else
	{
		#Server is an Exchange server, continue the health check

		#Custom object properties
		$serverObj = New-Object PSObject
		$serverObj | Add-Member NoteProperty -Name "Server" -Value $server
		
        #Skip Site attribute for Exchange 2003 servers
        if ($serverinfo.AdminDisplayVersion -like "Version 6.*")
		{
			$serverObj | Add-Member NoteProperty -Name "Site" -Value "n/a"
		}
        else
        {
		    $site = ($serverinfo.site.ToString()).Split("/")
		    $serverObj | Add-Member NoteProperty -Name "Site" -Value $site[-1]
        }
		
        #Null and n/a the rest, will be populated as script progresses
		$serverObj | Add-Member NoteProperty -Name "DNS" -Value $null
		$serverObj | Add-Member NoteProperty -Name "Ping" -Value $null
		$serverObj | Add-Member NoteProperty -Name "Uptime (hrs)" -Value $null
		$serverObj | Add-Member NoteProperty -Name "Version" -Value $null
		$serverObj | Add-Member NoteProperty -Name "Roles" -Value $null
		$serverObj | Add-Member NoteProperty -Name "Client Access Server Role Services" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "Hub Transport Server Role Services" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "Mailbox Server Role Services" -Value "n/a"
		#$serverObj | Add-Member NoteProperty -Name "Unified Messaging Server Role Services" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "Transport Queue" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "Queue Length" -Value "n/a"
		#$serverObj | Add-Member NoteProperty -Name "PF DBs Mounted" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "MB DBs Mounted" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "MAPI Test" -Value "n/a"

		#Check server name resolves in DNS
		if ($Log) {Write-Logfile $string30}
		Write-Host "DNS Check: " -NoNewline;
		try 
		{
			$fpp = @([System.Net.Dns]::GetHostByName($server).AddressList | Select-Object IPAddressToString -ExpandProperty IPAddressToString)
		}
		catch
		{
			Write-Host -ForegroundColor $warn $_.Exception.Message
			if ($Log) {Write-Logfile $_.Exception.Message}
			$fpp = $null
		}

		if ( $fpp -ne $null )
		{
			Write-Host -ForegroundColor $pass "Pass"
			if ($Log) {Write-Logfile $string31}
			$serverObj | Add-Member NoteProperty -Name "DNS" -Value "Pass" -Force

			#Is server online
			if ($Log) {Write-Logfile $string32}
			Write-Host "Ping Check: " -NoNewline; 
			
			$ping = $null
			try
			{
				$ping = Test-Connection $server -Quiet -ErrorAction Stop
			}
			catch
			{
				Write-Host -ForegroundColor $warn $_.Exception.Message
				if ($Log) {Write-Logfile $_.Exception.Message}
			}

			switch ($ping)
			{
				$true {
					Write-Host -ForegroundColor $pass "Pass"
					$serverObj | Add-Member NoteProperty -Name "Ping" -Value "Pass" -Force
					if ($Log) {Write-Logfile $string33}
					}
				default {
					Write-Host -ForegroundColor $fail "Fail"
					$serverObj | Add-Member NoteProperty -Name "Ping" -Value "Fail" -Force
					$serversummary += "$server - $string18"
					if ($Log) {Write-Logfile $string18}
					}
			}
			
			#Uptime check, even if ping fails
			if ($Log) {Write-Logfile $string34}
			[int]$uptime = $null
			#$laststart = $null
            $OS = $null
		
			try 
			{
				#$laststart = [System.Management.ManagementDateTimeconverter]::ToDateTime((Get-WmiObject -Class Win32_OperatingSystem -computername $server -ErrorAction Stop).LastBootUpTime)
                $OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $server -ErrorAction STOP
			}
			catch
			{
				Write-Host -ForegroundColor $warn $_.Exception.Message
				if ($Log) {Write-Logfile $_.Exception.Message}
			}
			
            Write-Host "Uptime (hrs): " -NoNewline

			if ($OS -eq $null)
			{
				[string]$uptime = $string17
				if ($Log) {Write-Logfile $string17}
				switch ($ping)
				{
                	$true {	$serversummary += "$server - $string17" }
					default { $serversummary += "$server - $string17" }
				}
			}
			else
			{
				$timespan = $OS.ConvertToDateTime($OS.LocalDateTime) – $OS.ConvertToDateTime($OS.LastBootUpTime)
				[int]$uptime = "{0:00}" -f $timespan.TotalHours
				Switch ($uptime -gt 23) {
				    $true { Write-Host -ForegroundColor $pass $uptime }
				    $false { Write-Host -ForegroundColor $warn $uptime; $serversummary += "$server - Uptime is less than 24 hours" }
				    default { Write-Host -ForegroundColor $warn $uptime; $serversummary += "$server - Uptime is less than 24 hours" }
			    }
			}

			if ($Log) {Write-Logfile "Uptime is $uptime hours"}

			$serverObj | Add-Member NoteProperty -Name "Uptime (hrs)" -Value $uptime -Force	
			
			if ($ping -or ($uptime -ne $string17))
			{
				#Determine the friendly version number
				$ExVer = $serverinfo.AdminDisplayVersion
				Write-Host "Server version: " -NoNewline;
				
				if ($ExVer -like "Version 6.*")
				{
					$version = "Exchange 2003"
				}
				
				if ($ExVer -like "Version 8.*")
				{
					$version = "Exchange 2007"
				}
				
				if ($ExVer -like "Version 14.*")
				{
					$version = "Exchange 2010"
				}
				
				if ($ExVer -like "Version 15.0*")
				{
					$version = "Exchange 2013"
				}

				if ($ExVer -like "Version 15.1*")
				{
					$version = "Exchange 2016"
				}

				if ($ExVer -like "Version 15.2*")
				{
					$version = "Exchange 2019"
				}
				
				Write-Host $version
				if ($Log) {Write-Logfile "Server is running $version"}
				$serverObj | Add-Member NoteProperty -Name "Version" -Value $version -Force
			
				if ($version -eq "Exchange 2003")
				{
					Write-Host $string12
					if ($Log) {Write-Logfile $string12}
				}

				#START - Exchange 2013/2010/2007 Health Checks
				if ($version -ne "Exchange 2003")
				{
					Write-Host "Roles:" $serverinfo.ServerRole
					if ($Log) {Write-Logfile "Server roles: $($serverinfo.ServerRole)"}
					$serverObj | Add-Member NoteProperty -Name "Roles" -Value $serverinfo.ServerRole -Force
					
					$fpsEdge = $serverinfo.IsEdgeServer		
					$fpsHub = $serverinfo.IsHubTransportServer
					$fpsCAS = $serverinfo.IsClientAccessServer
					$fpsMB = $serverinfo.IsMailboxServer

					#START - General Server Health Check
					#Skipping Edge Transports for the general health check, as firewalls usually get
					#in the way. If you want to include them, remove this If.
					if ($fpsEdge -ne $true)
					{
						#Service health is an array due to how multi-role servers return Test-ServiceHealth status
						if ($Log) {Write-Logfile $string35}
                        $servicehealth = @()
						$e15casservicehealth = @()
						try {
							$servicehealth = @(Test-ServiceHealth $server -ErrorAction Stop)
						}
						catch {
							#Workaround for Test-ServiceHealth problem with CAS-only Exchange 2013 servers
							#More info: http://exchangeserverpro.com/exchange-2013-test-servicehealth-error/
							if ($_.Exception.Message -like "*There are no Microsoft Exchange 2007 server roles installed*")
							{
								if ($Log) {Write-Logfile $string52}
								$e15casservicehealth = Test-E15CASServiceHealth($server)
							}
							else
							{
								$serversummary += "$server - $string4"
								Write-Host -ForegroundColor $warn $string4 ":" $_.Exception
								if ($Log) {Write-Logfile $_.Exception}
	                            $serverObj | Add-Member NoteProperty -Name "Client Access Server Role Services" -Value $string4 -Force
			                    $serverObj | Add-Member NoteProperty -Name "Hub Transport Server Role Services" -Value $string4 -Force
			                    $serverObj | Add-Member NoteProperty -Name "Mailbox Server Role Services" -Value $string4 -Force
			                    #$serverObj | Add-Member NoteProperty -Name "Unified Messaging Server Role Services" -Value $string4 -Force
							}
						}
							
						if ($servicehealth)
						{
							foreach($s in $servicehealth)
							{
								$roleName = $s.Role
								Write-Host $roleName "Services: " -NoNewline;
															
								switch ($s.RequiredServicesRunning)
								{
									$true {
										$svchealth = "Pass"
										Write-Host -ForegroundColor $pass "Pass"
										}
									$false {
										$svchealth = "Fail"
										Write-Host -ForegroundColor $fail "Fail"
										$serversummary += "$server - $rolename $string5"
										}
                                    default {
										$svchealth = "Warn"
										Write-Host -ForegroundColor $warn "Warning"
										$serversummary += "$server - $rolename $string5"
										}
								}

								switch ($s.Role)
								{
									$casrole { $serverinfoservices = "Client Access Server Role Services" }
									$htrole { $serverinfoservices = "Hub Transport Server Role Services" }
									$mbrole { $serverinfoservices = "Mailbox Server Role Services" }
									#$umrole { $serverinfoservices = "Unified Messaging Server Role Services" }
								}
								if ($Log) {Write-Logfile "$serverinfoservices status is $svchealth"}	
								$serverObj | Add-Member NoteProperty -Name $serverinfoservices -Value $svchealth -Force
							}
						}
						
						if ($e15casservicehealth)
						{
							$serverinfoservices = "Client Access Server Role Services"
							if ($Log) {Write-Logfile "$serverinfoservices status is $e15casservicehealth"}
							$serverObj | Add-Member NoteProperty -Name $serverinfoservices -Value $e15casservicehealth -Force
							Write-Host $serverinfoservices ": " -NoNewline;
							switch ($e15casservicehealth)
							{
								"Pass" { Write-Host -ForegroundColor $pass "Pass" }
								"Fail" { Write-Host -ForegroundColor $fail "Fail" }
							}
						}
					}
					#END - General Server Health Check

					#START - Hub Transport Server Check
					if ($fpsHub)
					{
						$q = $null
						if ($Log) {Write-Logfile $string36}
						Write-Host "Total Queue: " -NoNewline; 
						try {
							$q = Get-Queue -server $server -ErrorAction Stop
						}
						catch {
							$serversummary += "$server - $string6"
							Write-Host -ForegroundColor $warn $string6
							Write-Warning $_.Exception.Message
							if ($Log) {Write-Logfile $string6}
							if ($Log) {Write-Logfile $_.Exception.Message}
						}
						
						if ($q)
						{
							$qcount = $q | Measure-Object MessageCount -Sum
							[int]$qlength = $qcount.sum
							$serverObj | Add-Member NoteProperty -Name "Queue Length" -Value $qlength -Force
							if ($Log) {Write-Logfile "Queue length is $qlength"}
							if ($qlength -le $transportqueuewarn)
							{
								Write-Host -ForegroundColor $pass $qlength
								$serverObj | Add-Member NoteProperty -Name "Transport Queue" -Value "Pass ($qlength)" -Force
							}
							elseif ($qlength -gt $transportqueuewarn -and $qlength -lt $transportqueuehigh)
							{
								Write-Host -ForegroundColor $warn $qlength
                                $serversummary += "$server - Transport queue is above warning threshold" 
								$serverObj | Add-Member NoteProperty -Name "Transport Queue" -Value "Warn ($qlength)" -Force
							}
							else
							{
								Write-Host -ForegroundColor $fail $qlength
                                $serversummary += "$server - Transport queue is above high threshold"
								$serverObj | Add-Member NoteProperty -Name "Transport Queue" -Value "Fail ($qlength)" -Force
							}
						}
						else
						{
							$serverObj | Add-Member NoteProperty -Name "Transport Queue" -Value "Unknown" -Force
						}
					}
					#END - Hub Transport Server Check

					#START - Mailbox Server Check
					if ($fpsMB)
					{
						if ($Log) {Write-Logfile $string37}
						
						#Get the PF and MB databases
						#[array]$pfdbs = @(Get-PublicFolderDatabase -server $server -status -WarningAction SilentlyContinue)
						[array]$mbdbs = @(Get-MailboxDatabase -server $server -status | Where {$_.Recovery -ne $true})
                        
                        if ($version -ne "Exchange 2007")
                        {
						    [array]$activedbs = @(Get-MailboxDatabase -server $server -status | Where {$_.Recovery -ne $true -and $_.MountedOnServer -eq ($serverinfo.fqdn)})
                        }
                        else
                        {
                            [array]$activedbs = $mbdbs
                        }
						
						#START - Database Mount Check
						
						<#
                        #Check public folder databases
						if ($pfdbs.count -gt 0)
						{
							if ($Log) {Write-Logfile $string39}
							Write-Host "Public Folder databases mounted: " -NoNewline;
							[string]$pfdbstatus = "Pass"
							[array]$alertdbs = @()
							foreach ($db in $pfdbs)
							{
								if (($db.mounted) -ne $true)
								{
									$pfdbstatus = "Fail"
									$alertdbs += $db.name
								}
							}

							$serverObj | Add-Member NoteProperty -Name "PF DBs Mounted" -Value $pfdbstatus -Force
							if ($Log) {Write-Logfile "$string40 $pfdbstatus"}
							
							if ($alertdbs.count -eq 0)
							{
								Write-Host -ForegroundColor $pass $pfdbstatus
							}
							else
							{
								Write-Host -ForegroundColor $fail $pfdbstatus
								$serversummary += "$server - $string7"
								Write-Host "Offline databases:"
								foreach ($al in $alertdbs)
								{
									Write-Host -ForegroundColor $fail `t$al
								}
							}
						}
                        #>
						
						#Check mailbox databases
						if ($mbdbs.count -gt 0)
						{
							if ($Log) {Write-Logfile $string41}
						
							[string]$mbdbstatus = "Pass"
							[array]$alertdbs = @()

							Write-Host "Mailbox databases mounted: " -NoNewline;
							foreach ($db in $mbdbs)
							{
								if (($db.mounted) -ne $true)
								{
									$mbdbstatus = "Fail"
									$alertdbs += $db.name
								}
							}

							$serverObj | Add-Member NoteProperty -Name "MB DBs Mounted" -Value $mbdbstatus -Force
							if ($Log) {Write-Logfile "$string42 $mbdbstatus"}
							
							if ($alertdbs.count -eq 0)
							{
								Write-Host -ForegroundColor $pass $mbdbstatus
							}
							else
							{
								$serversummary += "$server - $string9"
								Write-Host -ForegroundColor $fail $mbdbstatus
								Write-Host $string43
								if ($Log) {Write-Logfile $string43}
								foreach ($al in $alertdbs)
								{
									Write-Host -ForegroundColor $fail `t$al
									if ($Log) {Write-Logfile "- $al"}
								}
							}
						}
						
						#END - Database Mount Check
						
						#START - MAPI Connectivity Test
						if ($activedbs.count -gt 0 -or $pfdbs.count -gt 0 -or $version -eq "Exchange 2007")
						{
							[string]$mapiresult = "Unknown"
							[array]$alertdbs = @()
							if ($Log) {Write-Logfile $string44}
							Write-Host "MAPI connectivity: " -NoNewline;
							foreach ($db in $mbdbs)
							{
								$mapistatus = Test-MapiConnectivity -Database $db.Identity -PerConnectionTimeout $mapitimeout
                                if ($mapistatus.Result.Value -eq $null)
                                {
                                    $mapiresult = $mapistatus.Result
                                }
                                else
                                {
                                    $mapiresult = $mapistatus.Result.Value
                                }
                                if (($mapiresult) -ne "Success")
								{
									$mapistatus = "Fail"
									$alertdbs += $db.name
								}
							}

							$serverObj | Add-Member NoteProperty -Name "MAPI Test" -Value  $mapiresult -Force
							if ($Log) {Write-Logfile "$string45  $mapiresult"}
							
							if ($alertdbs.count -eq 0)
							{
								Write-Host -ForegroundColor $pass  $mapiresult
							}
							else
							{
								$serversummary += "$server - $string10"
								Write-Host -ForegroundColor $fail  $mapiresult
								Write-Host $string46
								if ($Log) {Write-Logfile $string46}
								foreach ($al in $alertdbs)
								{
									Write-Host -ForegroundColor $fail `t$al
									if ($Log) {Write-Logfile "- $al"}
								}
							}
						}
						#END - MAPI Connectivity Test
						
						#START - Mail Flow Test
						if ($version -eq "Exchange 2007" -and $mbdbs.count -gt 0 -and $HasE15)
						{
							#Skip Exchange 2007 mail flow tests when run from Exchange 2013
							if ($Log) {Write-Logfile $string47}
							Write-Host "Mail flow test: Skipped"
							$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value $string51 -Force
							if ($Log) {Write-Logfile $string51}
						}
						elseif ($activedbs.count -gt 0 -and $HasE15)
						{
							if ($Log) {Write-Logfile $string47}
							Write-Host "Mail flow test: " -NoNewline;
							$e15mailflowresult = Test-E15MailFlow($Server)
							$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value $e15mailflowresult -Force
							if ($Log) {Write-Logfile "$string48 $e15mailflowresult"}
							
							if ($e15mailflowresult -eq $success)
							{
								Write-Host -ForegroundColor $pass $e15mailflowresult
								$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value "Pass" -Force
							}
							else
							{
								$serversummary += "$server - $string11"
								Write-Host -ForegroundColor $fail $e15mailflowresult
								$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value "Fail" -Force
							}
						}
						elseif ($activedbs.count -gt 0 -or ($version -eq "Exchange 2007" -and $mbdbs.count -gt 0))
						{
							$flow = $null
							$testmailflowresult = $null

							
							if ($Log) {Write-Logfile $string47}
							Write-Host "Mail flow test: " -NoNewline;
							try
							{
								$flow = Test-Mailflow $server -ErrorAction Stop
							}
							catch
							{
								$testmailflowresult = $_.Exception.Message
								if ($Log) {Write-Logfile $_.Exception.Message}
							}
							
							if ($flow)
							{
								$testmailflowresult = $flow.testmailflowresult
								if ($Log) {Write-Logfile "$string48 $testmailflowresult"}
							}

							if ($testmailflowresult -eq "Success" -or $testmailflowresult -eq $success)
							{
								Write-Host -ForegroundColor $pass $testmailflowresult
								$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value "Pass" -Force
							}
							else
							{
								$serversummary += "$server - $string11"
								Write-Host -ForegroundColor $fail $testmailflowresult
								$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value "Fail" -Force
							}
						}
						else
						{
							Write-Host "Mail flow test: No active mailbox databases"
							$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value $string49 -Force
							if ($Log) {Write-Logfile $string49}
						}
						#END - Mail Flow Test
					}
					#END - Mailbox Server Check

				}
				#END - Exchange 2013/2010/2007 Health Checks
				if ($Log) {Write-Logfile "$string50 $server"}
				$report = $report + $serverObj
			}
			else
			{
				#Server is not reachable and uptime could not be retrieved
				Write-Host -ForegroundColor $warn $string1
				if ($Log) {Write-Logfile $string1}
				$serversummary += "$server - $string1"
				$serverObj | Add-Member NoteProperty -Name "Ping" -Value "Fail" -Force
				if ($Log) {Write-Logfile "$string50 $server"}
				$report = $report + $serverObj
			}
		}
		else
		{
			Write-Host -ForegroundColor $Fail "Fail"
			Write-Host -ForegroundColor $warn $string13
			if ($Log) {Write-Logfile $string13}
			$serversummary += "$server - $string13"
			$serverObj | Add-Member NoteProperty -Name "DNS" -Value "Fail" -Force
			if ($Log) {Write-Logfile "$string50 $server"}
			$report = $report + $serverObj
		}
	}	
}
### End the Exchange Server health checks


### Begin DAG Health Report

#Check if -Server or -Serverlist parameter was used, and skip if it was
if (!($NoDAG))
{
	if ($Log) {Write-Logfile $string60}
	Write-Verbose "Retrieving Database Availability Groups"

	#Get all DAGs
	$tmpdags = @(Get-DatabaseAvailabilityGroup)
	$tmpstring = "$($tmpdags.count) DAGs found"
	Write-Verbose $tmpstring
	if ($Log) {Write-Logfile $tmpstring}

	#Remove DAGs in ignorelist
	foreach ($tmpdag in $tmpdags)
	{
		if (!($fpgnorelist -icontains $tmpdag.name))
		{
			$dags += $tmpdag
		}
	}

	$tmpstring = "$($dags.count) DAGs will be checked"
	Write-Verbose $tmpstring
	if ($Log) {Write-Logfile $tmpstring}

	if ($Log) {Write-Logfile $string68}
	if ($Log) {
		foreach ($dag in $dags)
		{
			Write-Logfile "- $dag"
		}
	}
}

if ($($dags.count) -gt 0)
{
	foreach ($dag in $dags)
	{
		
		#Strings for use in the HTML report/email
		$dagsummaryintro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Health Summary:</p>"
		$dagdetailintro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Health Details:</p>"
		$dagmemberintro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Member Health:</p>"

		$dagdbcopyReport = @()		#Database copy health report
		$dagciReport = @()			#Content Index health report
		$dagmemberReport = @()		#DAG member server health report
		$dagdatabaseSummary = @()	#Database health summary report
		$dagdatabases = @()			#Array of databases in the DAG
		
		$tmpstring = "---- Processing DAG $($dag.Name)"
		Write-Verbose $tmpstring
		if ($Log) {Write-Logfile $tmpstring}
		
		$dagmembers = @($dag | Select-Object -ExpandProperty Servers | Sort-Object Name)
		$tmpstring = "$($dagmembers.count) DAG members found"
		Write-Verbose $tmpstring
		if ($Log) {Write-Logfile $tmpstring}
		
		#Get all databases in the DAG
        if ($HasE15)
        {
		    $tmpdatabases = @(Get-MailboxDatabase -Status -IncludePreExchange2013 | Where-Object {$_.MasterServerOrAvailabilityGroup -eq $dag.Name} | Sort-Object Name)
        }
        else
        {
		    $tmpdatabases = @(Get-MailboxDatabase -Status | Where-Object {$_.MasterServerOrAvailabilityGroup -eq $dag.Name} | Sort-Object Name)
        }

		foreach ($tmpdatabase in $tmpdatabases)
		{
			if (!($fpgnorelist -icontains $tmpdatabase.name))
			{
				$dagdatabases += $tmpdatabase
			}
		}
				
		$tmpstring = "$($dagdatabases.count) DAG databases will be checked"
		Write-Verbose $tmpstring
		if ($Log) {Write-Logfile $tmpstring}

		if ($Log) {Write-Logfile $string69}
		if ($Log) {
			foreach ($database in $dagdatabases)
			{
				Write-Logfile "- $database"
			}
		}
		
		foreach ($database in $dagdatabases)
		{
			$tmpstring = "---- Processing database $database"
			Write-Verbose $tmpstring
			if ($Log) {Write-Logfile $tmpstring}

			#Custom object for Database
			$objectHash = @{
				"Database" = $database.Identity
				"Mounted on" = "Unknown"
				#"Preference" = $null
				"Total Copies" = $null
				"Healthy Copies" = $null
				"Unhealthy Copies" = $null
				"Healthy Queues" = $null
				"Unhealthy Queues" = $null
				"Lagged Queues" = $null
				#"Healthy Indexes" = $null
				#"Unhealthy Indexes" = $null
				}
			$databaseObj = New-Object PSObject -Property $objectHash

			$dbcopystatus = @($database | Get-MailboxDatabaseCopyStatus)
			$tmpstring = "$database has $($dbcopystatus.Count) copies"
			Write-Verbose $tmpstring
			if ($Log) {Write-Logfile $tmpstring}
			
			foreach ($dbcopy in $dbcopystatus)
			{
				#Custom object for DB copy
				$objectHash = @{
					"Database Copy" = $dbcopy.Identity
					"Database Name" = $dbcopy.DatabaseName
					"Mailbox Server" = $null
					#"Activation Preference" = $null
					"Status" = $null
					"Copy Queue" = $null
					"Replay Queue" = $null
					"Replay Lagged" = $null
					"Truncation Lagged" = $null
					"Content Index" = $null
					}
				$dbcopyObj = New-Object PSObject -Property $objectHash
				
				$tmpstring = "Database Copy: $($dbcopy.Identity)"
				Write-Verbose $tmpstring
				if ($Log) {Write-Logfile $tmpstring}
				
				$mailboxserver = $dbcopy.MailboxServer
				$tmpstring = "Server: $mailboxserver"
				Write-Verbose $tmpstring
				if ($Log) {Write-Logfile $tmpstring}

				#$pref = ($database | Select-Object -ExpandProperty ActivationPreference | Where-Object {$_.Key -eq $mailboxserver}).Value
				#$tmpstring = "Activation Preference: $pref"
				#Write-Verbose $tmpstring
				if ($Log) {Write-Logfile $tmpstring}

				$copystatus = $dbcopy.Status
				$tmpstring = "Status: $copystatus"
				Write-Verbose $tmpstring
				if ($Log) {Write-Logfile $tmpstring}
				
				[int]$copyqueuelength = $dbcopy.CopyQueueLength
				$tmpstring = "Copy Queue: $copyqueuelength"
				Write-Verbose $tmpstring
				if ($Log) {Write-Logfile $tmpstring}
				
				[int]$replayqueuelength = $dbcopy.ReplayQueueLength
				$tmpstring = "Replay Queue: $replayqueuelength"
				Write-Verbose $tmpstring
				if ($Log) {Write-Logfile $tmpstring}
				
				if ($($dbcopy.ContentIndexErrorMessage -match "is disabled in Active Directory"))
                {
                    $contentindexstate = "Disabled"
                }
                else
                {
                    $contentindexstate = $dbcopy.ContentIndexState
                }
				$tmpstring = "Content Index: $contentindexstate"
				Write-Verbose $tmpstring
				if ($Log) {Write-Logfile $tmpstring}				

				#Checking whether this is a replay lagged copy
				$replaylagcopies = @($database | Select -ExpandProperty ReplayLagTimes | Where-Object {$_.Value -gt 0})
				if ($($replaylagcopies.count) -gt 0)
	            {
	                [bool]$replaylag = $false
	                foreach ($replaylagcopy in $replaylagcopies)
				    {
					    if ($replaylagcopy.Key -eq $mailboxserver)
					    {
						    $tmpstring = "$database is replay lagged on $mailboxserver"
							Write-Verbose $tmpstring
							if ($Log) {Write-Logfile $tmpstring}
						    [bool]$replaylag = $true
					    }
				    }
	            }
	            else
				{
				   [bool]$replaylag = $false
				}
	            $tmpstring = "Replay lag is $replaylag"
				Write-Verbose $tmpstring
				if ($Log) {Write-Logfile $tmpstring}				
						
				#Checking for truncation lagged copies
				$truncationlagcopies = @($database | Select -ExpandProperty TruncationLagTimes | Where-Object {$_.Value -gt 0})
				if ($($truncationlagcopies.count) -gt 0)
	            {
	                [bool]$truncatelag = $false
	                foreach ($truncationlagcopy in $truncationlagcopies)
				    {
					    if ($truncationlagcopy.Key -eq $mailboxserver)
					    {
						    $tmpstring = "$database is truncate lagged on $mailboxserver"
							Write-Verbose $tmpstring
							if ($Log) {Write-Logfile $tmpstring}							
							[bool]$truncatelag = $true
					    }
				    }
	            }
	            else
				{
				   [bool]$truncatelag = $false
				}
	            $tmpstring = "Truncation lag is $truncatelag"
				Write-Verbose $tmpstring
				if ($Log) {Write-Logfile $tmpstring}
				
				$dbcopyObj | Add-Member NoteProperty -Name "Mailbox Server" -Value $mailboxserver -Force
				#$dbcopyObj | Add-Member NoteProperty -Name "Activation Preference" -Value $pref -Force
				$dbcopyObj | Add-Member NoteProperty -Name "Status" -Value $copystatus -Force
				$dbcopyObj | Add-Member NoteProperty -Name "Copy Queue" -Value $copyqueuelength -Force
				$dbcopyObj | Add-Member NoteProperty -Name "Replay Queue" -Value $replayqueuelength -Force
				$dbcopyObj | Add-Member NoteProperty -Name "Replay Lagged" -Value $replaylag -Force
				$dbcopyObj | Add-Member NoteProperty -Name "Truncation Lagged" -Value $truncatelag -Force
				$dbcopyObj | Add-Member NoteProperty -Name "Content Index" -Value $contentindexstate -Force
				
				$dagdbcopyReport += $dbcopyObj
			}
		
			$copies = @($dagdbcopyReport | Where-Object { ($_."Database Name" -eq $database) })
		
			$mountedOn = ($copies | Where-Object { ($_.Status -eq "Mounted") })."Mailbox Server"
			if ($mountedOn)
			{
				$databaseObj | Add-Member NoteProperty -Name "Mounted on" -Value $mountedOn -Force
			}
		
			#$activationPref = ($copies | Where-Object { ($_.Status -eq "Mounted") })."Activation Preference"
			#$databaseObj | Add-Member NoteProperty -Name "Preference" -Value $activationPref -Force

			$totalcopies = $copies.count
			$databaseObj | Add-Member NoteProperty -Name "Total Copies" -Value $totalcopies -Force
		
			$healthycopies = @($copies | Where-Object { (($_.Status -eq "Mounted") -or ($_.Status -eq "Healthy")) }).Count
			$databaseObj | Add-Member NoteProperty -Name "Healthy Copies" -Value $healthycopies -Force
			
			$unhealthycopies = @($copies | Where-Object { (($_.Status -ne "Mounted") -and ($_.Status -ne "Healthy")) }).Count
			$databaseObj | Add-Member NoteProperty -Name "Unhealthy Copies" -Value $unhealthycopies -Force

			$healthyqueues  = @($copies | Where-Object { (($_."Copy Queue" -lt $replqueuewarning) -and (($_."Replay Queue" -lt $replqueuewarning)) -and ($_."Replay Lagged" -eq $false)) }).Count
	        $databaseObj | Add-Member NoteProperty -Name "Healthy Queues" -Value $healthyqueues -Force

			$unhealthyqueues = @($copies | Where-Object { (($_."Copy Queue" -ge $replqueuewarning) -or (($_."Replay Queue" -ge $replqueuewarning) -and ($_."Replay Lagged" -eq $false))) }).Count
			$databaseObj | Add-Member NoteProperty -Name "Unhealthy Queues" -Value $unhealthyqueues -Force

			#$laggedqueues = @($copies | Where-Object { ($_."Replay Lagged" -eq $true) -or ($_."Truncation Lagged" -eq $true) }).Count
			#$databaseObj | Add-Member NoteProperty -Name "Lagged Queues" -Value $laggedqueues -Force

			#$healthyindexes = @($copies | Where-Object { ($_."Content Index" -eq "Healthy" -or $_."Content Index" -eq "Disabled") }).Count
			#$databaseObj | Add-Member NoteProperty -Name "Healthy Indexes" -Value $healthyindexes -Force
			
			#$unhealthyindexes = @($copies | Where-Object { ($_."Content Index" -ne "Healthy" -and $_."Content Index" -ne "Disabled") }).Count
			#$databaseObj | Add-Member NoteProperty -Name "Unhealthy Indexes" -Value $unhealthyindexes -Force
			
			$dagdatabaseSummary += $databaseObj
		
		}
		
		#Get Test-Replication Health results for each DAG member
		foreach ($dagmember in $dagmembers)
		{
            $replicationhealth = $null

            $replicationhealthitems = @{ClusterService = $null
                                        ReplayService = $null
                                        ActiveManager = $null
                                        TasksRpcListener = $null
                                        TcpListener = $null
                                        ServerLocatorService = $null
                                        DagMembersUp = $null
                                        ClusterNetwork = $null
                                        QuorumGroup = $null
                                        FileShareQuorum = $null
                                        DatabaseRedundancy = $null
                                        DatabaseAvailability = $null
                                        DBCopySuspended = $null
                                        DBCopyFailed = $null
                                        DBInitializing = $null
                                        DBDisconnected = $null
                                        DBLogCopyKeepingUp = $null
                                        DBLogReplayKeepingUp = $null
                                        }

			$memberObj = New-Object PSObject -Property $replicationhealthitems
			$memberObj | Add-Member NoteProperty -Name "Server" -Value $dagmember
		
			$tmpstring = "---- Checking replication health for $dagmember"
			Write-Verbose $tmpstring
			if ($Log) {Write-Logfile $tmpstring}
			
			try
            {
                $replicationhealth = $dagmember | Invoke-Command {Test-ReplicationHealth -ErrorAction STOP} 
            }
            catch
            {
		        if ($Log) {Write-Logfile "Using e15 replication health test workaround"}
                $replicationhealth = Test-e15ReplicationHealth $dagmember
            }
			
	        foreach ($healthitem in $replicationhealth)
	        {
                if ($($healthitem.Result) -eq $null)
                {
                    $healthitemresult = "n/a"
                }
                else
                {
                    $healthitemresult = $($healthitem.Result)
                }
                $tmpstring = "$($healthitem.Check) $healthitemresult"
		        Write-Verbose $tmpstring
		        if ($Log) {Write-Logfile $tmpstring}
		        $memberObj | Add-Member NoteProperty -Name $($healthitem.Check) -Value $healthitemresult -Force
	        }
			$dagmemberReport += $memberObj
		}

		
		#Generate the HTML from the DAG health checks
		if ($SendEmail -or $ReportFile)
		{
		
			####Begin Summary Table HTML
			$dagdatabaseSummaryHtml = $null
			#Begin Summary table HTML header
			$htmltableheader = "<p>
							<table>
							<tr>
							<th>Database</th>
							<th>Mounted on</th>
							<th>Total Copies</th>
							<th>Healthy Copies</th>
							<th>Unhealthy Copies</th>
							<th>Healthy Queues</th>
							<th>Unhealthy Queues</th>
							</tr>"

			$dagdatabaseSummaryHtml += $htmltableheader
			#End Summary table HTML header
                        #<th>Preference</th>
                        #<th>Lagged Queues</th>
			
			#Begin Summary table HTML rows
			foreach ($line in $dagdatabaseSummary)
			{
				$htmltablerow = "<tr>"
				$htmltablerow += "<td><strong>$($line.Database)</strong></td>"
				
				#Warn if mounted server is still unknown
				switch ($($line."Mounted on"))
				{
					"Unknown" {
						$htmltablerow += "<td class=""warn"">$($line."Mounted on")</td>"
						$dagsummary += "$($line.Database) - $string61"
						}
					default { $htmltablerow += "<td>$($line."Mounted on")</td>" }
				}
				
				#Warn if DB is mounted on a server that is not Activation Preference 1
				<#
                if ($($line.Preference) -gt 1)
				{
					$htmltablerow += "<td class=""warn"">$($line.Preference)</td>"
					$dagsummary += "$($line.Database) - $string62 $($line.Preference)"
				}
				else
				{
					$htmltablerow += "<td class=""pass"">$($line.Preference)</td>"
				}
                #>
				
				$htmltablerow += "<td>$($line."Total Copies")</td>"
				
				#Show as info if health copies is 1 but total copies also 1,
	            #Warn if healthy copies is 1, Fail if 0
				switch ($($line."Healthy Copies"))
				{	
					0 {$htmltablerow += "<td class=""fail"">$($line."Healthy Copies")</td>"}
					1 {
						if ($($line."Total Copies") -eq $($line."Healthy Copies"))
						{
							$htmltablerow += "<td class=""info"">$($line."Healthy Copies")</td>"
						}
						else
						{
							$htmltablerow += "<td class=""warn"">$($line."Healthy Copies")</td>"
						}
					  }
					default {$htmltablerow += "<td class=""pass"">$($line."Healthy Copies")</td>"}
				}

				#Warn if unhealthy copies is 1, fail if more than 1
				switch ($($line."Unhealthy Copies"))
				{
					0 {	$htmltablerow += "<td class=""pass"">$($line."Unhealthy Copies")</td>" }
					1 {
						$htmltablerow += "<td class=""warn"">$($line."Unhealthy Copies")</td>"
						$dagsummary += "$($line.Database) - $string63 $($line."Unhealthy Copies") $string65 $($line."Total Copies") $string66"
						}
					default {
						$htmltablerow += "<td class=""fail"">$($line."Unhealthy Copies")</td>"
						$dagsummary += "$($line.Database) - $string63 $($line."Unhealthy Copies") $string65 $($line."Total Copies") $string66"
						}
				}

				#Warn if healthy queues + lagged queues is less than total copies
				#Fail if no healthy queues
				if ($($line."Total Copies") -eq ($($line."Healthy Queues") + $($line."Lagged Queues")))
				{
					$htmltablerow += "<td class=""pass"">$($line."Healthy Queues")</td>"
				}
				else
				{
					$dagsummary += "$($line.Database) - $string64 $($line."Healthy Queues") $string65 $($line."Total Copies") $string66"
					switch ($($line."Healthy Queues"))
					{
						0 {	$htmltablerow += "<td class=""fail"">$($line."Healthy Queues")</td>" }
						default { $htmltablerow += "<td class=""warn"">$($line."Healthy Queues")</td>" }
					}
				}
				
				#Fail if unhealthy queues = total queues
				#Warn if more than one unhealthy queue
				if ($($line."Total Queues") -eq $($line."Unhealthy Queues"))
				{
					$htmltablerow += "<td class=""fail"">$($line."Unhealthy Queues")</td>"
				}
				else
				{
					switch ($($line."Unhealthy Queues"))
					{
						0 { $htmltablerow += "<td class=""pass"">$($line."Unhealthy Queues")</td>" }
						default { $htmltablerow += "<td class=""warn"">$($line."Unhealthy Queues")</td>" }
					}
				}
				
				#Info for lagged queues
				<#
                switch ($($line."Lagged Queues"))
				{
					0 { $htmltablerow += "<td>$($line."Lagged Queues")</td>" }
					default { $htmltablerow += "<td class=""info"">$($line."Lagged Queues")</td>" }
				}
                
                #>

				########################################################################################################
                # Section below has been commented out - index checks no longer needed for Exch 2019. 6/23/23 HB
                ########################################################################################################
				#Pass if healthy indexes = total copies
				#Warn if healthy indexes less than total copies
				#Fail if healthy indexes = 0
				#if ($($line."Total Copies") -eq $($line."Healthy Indexes"))
				#{
				#	$htmltablerow += "<td class=""pass"">$($line."Healthy Indexes")</td>"
				#}
				#else
				#{
				#	$dagsummary += "$($line.Database) - $string67 $($line."Unhealthy Indexes") $string65 $($line."Total Copies") $string66"
				#	switch ($($line."Healthy Indexes"))
				#	{
				#		0 { $htmltablerow += "<td class=""fail"">$($line."Healthy Indexes")</td>" }
				#		default { $htmltablerow += "<td class=""warn"">$($line."Healthy Indexes")</td>" }
				#	}
				#}
				
				#Fail if unhealthy indexes = total copies
				#Warn if unhealthy indexes 1 or more
				#Pass if unhealthy indexes = 0
				#if ($($line."Total Copies") -eq $($line."Unhealthy Indexes"))
				#{
				#	$htmltablerow += "<td class=""fail"">$($line."Unhealthy Indexes")</td>"
				#}
				#else
				#{
				#	switch ($($line."Unhealthy Indexes"))
				#	{
				#		0 { $htmltablerow += "<td class=""pass"">$($line."Unhealthy Indexes")</td>" }
				#		default { $htmltablerow += "<td class=""warn"">$($line."Unhealthy Indexes")</td>" }
				#	}
				#}
                ########################################################################################################
				
				$htmltablerow += "</tr>"
				$dagdatabaseSummaryHtml += $htmltablerow
			}
			$dagdatabaseSummaryHtml += "</table>
									</p>"
			#End Summary table HTML rows
			####End Summary Table HTML

			####Begin Detail Table HTML
			$databasedetailsHtml = $null
			#Begin Detail table HTML header
			$htmltableheader = "<p>
							<table>
							<tr>
							<th>Database Copy</th>
							<th>Database Name</th>
							<th>Mailbox Server</th>
							<th>Status</th>
							<th>Copy Queue</th>
							<th>Replay Queue</th>
							<th>Replay Lagged</th>
							<th>Truncation Lagged</th>
							<th>Content Index</th>
							</tr>"

			$databasedetailsHtml += $htmltableheader
			#End Detail table HTML header
                        #<th>Activation Preference</th>
			
			#Begin Detail table HTML rows
			foreach ($line in $dagdbcopyReport)
			{
				$htmltablerow = "<tr>"
				$htmltablerow += "<td><strong>$($line."Database Copy")</strong></td>"
				$htmltablerow += "<td>$($line."Database Name")</td>"
				$htmltablerow += "<td>$($line."Mailbox Server")</td>"
				#$htmltablerow += "<td>$($line."Activation Preference")</td>"
				
				Switch ($($line."Status"))
				{
					"Healthy" { $htmltablerow += "<td class=""pass"">$($line."Status")</td>" }
					"Mounted" { $htmltablerow += "<td class=""pass"">$($line."Status")</td>" }
					"Failed" { $htmltablerow += "<td class=""fail"">$($line."Status")</td>" }
					"FailedAndSuspended" { $htmltablerow += "<td class=""fail"">$($line."Status")</td>" }
					"ServiceDown" { $htmltablerow += "<td class=""fail"">$($line."Status")</td>" }
					"Dismounted" { $htmltablerow += "<td class=""fail"">$($line."Status")</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line."Status")</td>" }
				}
				
				if ($($line."Copy Queue") -lt $replqueuewarning)
				{
					$htmltablerow += "<td class=""pass"">$($line."Copy Queue")</td>"
				}
				else
				{
					$htmltablerow += "<td class=""warn"">$($line."Copy Queue")</td>"
				}
				
				if (($($line."Replay Queue") -lt $replqueuewarning) -or ($($line."Replay Lagged") -eq $true))
				{
					$htmltablerow += "<td class=""pass"">$($line."Replay Queue")</td>"
				}
				else
				{
					$htmltablerow += "<td class=""warn"">$($line."Replay Queue")</td>"
				}
				

				Switch ($($line."Replay Lagged"))
				{
					$true { $htmltablerow += "<td class=""info"">$($line."Replay Lagged")</td>" }
					default { $htmltablerow += "<td>$($line."Replay Lagged")</td>" }
				}

				Switch ($($line."Truncation Lagged"))
				{
					$true { $htmltablerow += "<td class=""info"">$($line."Truncation Lagged")</td>" }
					default { $htmltablerow += "<td>$($line."Truncation Lagged")</td>" }
				}
				
				Switch ($($line."Content Index"))
				{
					"Healthy" { $htmltablerow += "<td class=""pass"">$($line."Content Index")</td>" }
                    "Disabled" { $htmltablerow += "<td class=""info"">$($line."Content Index")</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line."Content Index")</td>" }
				}
				
				$htmltablerow += "</tr>"
				$databasedetailsHtml += $htmltablerow
			}
			$databasedetailsHtml += "</table>
									</p>"
			#End Detail table HTML rows
			####End Detail Table HTML
			
			
			####Begin Member Table HTML
			$dagmemberHtml = $null
			#Begin Member table HTML header
			$htmltableheader = "<p>
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
								<th>File Share Quorum</th>
								<th>Database Redundancy</th>
								<th>Database Availability</th>
								<th>DB Copy Suspended</th>
								<th>DB Copy Failed</th>
								<th>DB Initializing</th>
								<th>DB Disconnected</th>
								<th>DB Log Copy Keeping Up</th>
								<th>DB Log Replay Keeping Up</th>
								</tr>"
			
			$dagmemberHtml += $htmltableheader
			#End Member table HTML header
			
			#Begin Member table HTML rows
			foreach ($line in $dagmemberReport)
			{
				$htmltablerow = "<tr>"
				$htmltablerow += "<td><strong>$($line."Server")</strong></td>"
				$htmltablerow += (New-DAGMemberHTMLTableCell "ClusterService")
				$htmltablerow += (New-DAGMemberHTMLTableCell "ReplayService")
				$htmltablerow += (New-DAGMemberHTMLTableCell "ActiveManager")
				$htmltablerow += (New-DAGMemberHTMLTableCell "TasksRPCListener")
				$htmltablerow += (New-DAGMemberHTMLTableCell "TCPListener")
				$htmltablerow += (New-DAGMemberHTMLTableCell "ServerLocatorService")
				$htmltablerow += (New-DAGMemberHTMLTableCell "DAGMembersUp")
				$htmltablerow += (New-DAGMemberHTMLTableCell "ClusterNetwork")
				$htmltablerow += (New-DAGMemberHTMLTableCell "QuorumGroup")
				$htmltablerow += (New-DAGMemberHTMLTableCell "FileShareQuorum")
				$htmltablerow += (New-DAGMemberHTMLTableCell "DatabaseRedundancy")
				$htmltablerow += (New-DAGMemberHTMLTableCell "DatabaseAvailability")
				$htmltablerow += (New-DAGMemberHTMLTableCell "DBCopySuspended")
				$htmltablerow += (New-DAGMemberHTMLTableCell "DBCopyFailed")
				$htmltablerow += (New-DAGMemberHTMLTableCell "DBInitializing")
				$htmltablerow += (New-DAGMemberHTMLTableCell "DBDisconnected")
				$htmltablerow += (New-DAGMemberHTMLTableCell "DBLogCopyKeepingUp")
				$htmltablerow += (New-DAGMemberHTMLTableCell "DBLogReplayKeepingUp")
				$htmltablerow += "</tr>"
				$dagmemberHtml += $htmltablerow
			}
			$dagmemberHtml += "</table>
			</p>"
		}
		
		#Output the report objects to console, and optionally to email and HTML file
		#Forcing table format for console output due to issue with multiple output
		#objects that have different layouts

		#Write-Host "---- Database Copy Health Summary ----"
		#$dagdatabaseSummary | ft
				
		#Write-Host "---- Database Copy Health Details ----"
		#$dagdbcopyReport | ft
		
		#Write-Host "`r`n---- Server Test-Replication Report ----`r`n"
		#$dagmemberReport | ft
		
		if ($SendEmail -or $ReportFile)
		{
			$dagreporthtml = $dagsummaryintro + $dagdatabaseSummaryHtml + $dagdetailintro + $databasedetailsHtml + $dagmemberintro + $dagmemberHtml
			$dagreportbody += $dagreporthtml
		}
		
	}
}
else
{
	$tmpstring = "No DAGs found"
	if ($Log) {Write-LogFile $tmpstring}
	Write-Verbose $tmpstring
	$dagreporthtml = "<p>No database availability groups found.</p>"
}
### End DAG Health Report

Write-Host $string16
### Begin report generation
if ($ReportMode -or $SendEmail)
{
	#Get report generation timestamp
	$reportime = Get-Date

	#Create HTML Report
	#Common HTML head and styles
	$htmlhead="<html>
				<style>
				BODY{font-family: Arial; font-size: 8pt;}
				H1{font-size: 16px;}
				H2{font-size: 14px;}
				H3{font-size: 12px;}
				TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
				TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
				TD{border: 1px solid black; padding: 5px; }
				td.pass{background: #7FFF00;}
				td.warn{background: #FFE600;}
				td.fail{background: #FF0000; color: #ffffff;}
				td.info{background: #85D4FF;}
				</style>
				<body>
				<h1 align=""center"">Exchange Server Health Check Report</h1>
				<h3 align=""center"">Generated: $reportime</h3>"

	#Check if the server summary has 1 or more entries
	if ($($serversummary.count) -gt 0)
	{
		#Set alert flag to true
		$alerts = $true
	
		#Generate the HTML
		$serversummaryhtml = "<h3>Exchange Server Health Check Summary</h3>
						<p>The following server errors and warnings were detected.</p>
						<p>
						<ul>"
		foreach ($reportline in $serversummary)
		{
			$serversummaryhtml +="<li>$reportline</li>"
		}
		$serversummaryhtml += "</ul></p>"
		$alerts = $true
	}
	else
	{
		#Generate the HTML to show no alerts
		$serversummaryhtml = "<h3>Exchange Server Health Check Summary</h3>
						<p>No Exchange server health errors or warnings.</p>"
	}
	
	#Check if the DAG summary has 1 or more entries
	if ($($dagsummary.count) -gt 0)
	{
		#Set alert flag to true
		$alerts = $true
	
		#Generate the HTML
		$dagsummaryhtml = "<h3>Database Availability Group Health Check Summary</h3>
						<p>The following DAG errors and warnings were detected.</p>
						<p>
						<ul>"
		foreach ($reportline in $dagsummary)
		{
			$dagsummaryhtml +="<li>$reportline</li>"
		}
		$dagsummaryhtml += "</ul></p>"
		$alerts = $true
	}
	else
	{
		#Generate the HTML to show no alerts
		$dagsummaryhtml = "<h3>Database Availability Group Health Check Summary</h3>
						<p>No Exchange DAG errors or warnings.</p>"
	}


	#Exchange Server Health Report Table Header
	$htmltableheader = "<h3>Exchange Server Health</h3>
						<p>
						<table>
						<tr>
						<th>Server</th>
						<th>Site</th>
						<th>Roles</th>
						<th>Version</th>
						<th>DNS</th>
						<th>Ping</th>
						<th>Uptime (hrs)</th>
						<th>Client Access Server Role Services</th>
						<th>Hub Transport Server Role Services</th>
						<th>Mailbox Server Role Services</th>
						<th>Transport Queue</th>
						<th>MB DBs Mounted</th>
						<th>MAPI Test</th>
						<th>Mail Flow Test</th>
						</tr>"

	#Exchange Server Health Report Table
                #<th>Unified Messaging Server Role Services</th>
                #<th>PF DBs Mounted</th>
	$serverhealthhtmltable = $serverhealthhtmltable + $htmltableheader					
						
	foreach ($reportline in $report)
	{
		$htmltablerow = "<tr>"
		$htmltablerow += "<td>$($reportline.server)</td>"
		$htmltablerow += "<td>$($reportline.site)</td>"
		$htmltablerow += "<td>$($reportline.roles)</td>"
		$htmltablerow += "<td>$($reportline.version)</td>"					
		$htmltablerow += (New-ServerHealthHTMLTableCell "dns")
		$htmltablerow += (New-ServerHealthHTMLTableCell "ping")
		
		if ($($reportline."uptime (hrs)") -eq "Access Denied")
		{
			$htmltablerow += "<td class=""warn"">Access Denied</td>"		
		}
        elseif ($($reportline."uptime (hrs)") -eq $string17)
        {
            $htmltablerow += "<td class=""warn"">$string17</td>"
        }
		else
		{
			$hours = [int]$($reportline."uptime (hrs)")
			if ($hours -le 24)
			{
				$htmltablerow += "<td class=""warn"">$hours</td>"
			}
			else
			{
				$htmltablerow += "<td class=""pass"">$hours</td>"
			}
		}

		$htmltablerow += (New-ServerHealthHTMLTableCell "Client Access Server Role Services")
		$htmltablerow += (New-ServerHealthHTMLTableCell "Hub Transport Server Role Services")
		$htmltablerow += (New-ServerHealthHTMLTableCell "Mailbox Server Role Services")
		#$htmltablerow += (New-ServerHealthHTMLTableCell "Unified Messaging Server Role Services")
		#$htmltablerow += (New-ServerHealthHTMLTableCell "Transport Queue")
        if ($($reportline."Transport Queue") -match "Pass")
        {
            $htmltablerow += "<td class=""pass"">$($reportline."Transport Queue")</td>"
        }
        elseif ($($reportline."Transport Queue") -match "Warn")
        {
            $htmltablerow += "<td class=""warn"">$($reportline."Transport Queue")</td>"
        }
        elseif ($($reportline."Transport Queue") -match "Fail")
        {
            $htmltablerow += "<td class=""fail"">$($reportline."Transport Queue")</td>"
        }
        elseif ($($reportline."Transport Queue") -eq "n/a")
        {
            $htmltablerow += "<td>$($reportline."Transport Queue")</td>"
        }
        else
        {
            $htmltablerow += "<td class=""warn"">$($reportline."Transport Queue")</td>"
        }
		#$htmltablerow += (New-ServerHealthHTMLTableCell "PF DBs Mounted")
		$htmltablerow += (New-ServerHealthHTMLTableCell "MB DBs Mounted")
		$htmltablerow += (New-ServerHealthHTMLTableCell "MAPI Test")
		$htmltablerow += (New-ServerHealthHTMLTableCell "Mail Flow Test")
		$htmltablerow += "</tr>"
		
		$serverhealthhtmltable = $serverhealthhtmltable + $htmltablerow
	}

	$serverhealthhtmltable = $serverhealthhtmltable + "</table></p>"

	$htmltail = "</body>
				</html>"

	$htmlreport = $htmlhead + $serversummaryhtml + $dagsummaryhtml + $serverhealthhtmltable + $dagreportbody + $htmltail
	
	if ($ReportMode -or $ReportFile)
	{
		$htmlreport | Out-File $ReportFile -Encoding UTF8
	}

	if ($SendEmail)
	{
		if ($alerts -eq $false -and $AlertsOnly -eq $true)
		{
			#Do not send email message
			Write-Host $string19
			if ($Log) {Write-Logfile $string19}
		}
		else
		{
			#Send email message
			Write-Host $string14
			#Invoke-Command -ComputerName N490MBX04 -ScriptBlock { 
                Send-MailMessage @smtpsettings -Body $htmlreport -BodyAsHtml 
            #}
		}
	}
}
### End report generation


Write-Host $string15
if ($Log) {Write-Logfile $string15}

Remove-PSSession $RemoteSession

$logpath = "\\newton\admin\SysAdmin Powershell Module\Scripts\BattleRythm\Logs\Log_BattleRhythm.csv"
Add-Content -Path $logpath -Value "`Test-ExchangeServerHealth,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami)"

}

#########################################################





#########################################################
#                                                       #
# These Functions are for Support Purposes              #
#                                                       #
#########################################################
Function BattleRhythm {

<#
 .Synopsis
  Automates any type of monthly script. Includes last ran date. 

 .Description
  Allows admins to quickly run monthly scripts

 .Example
   
   
   

#>

[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
[void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')
[void][System.Reflection.Assembly]::LoadWithPartialName('PresentationFrameWork')
#Add-Type -AssemblyName PresentationFramework


#----- GUIDE FOR ADDING NEW SCRIPTS TO THE BATTLE RHYTHM -----#
<# 

1. Create a variable to check the last run date - ex. $LastGetStaleUsers = (($lastrun | Where-Object {$_.FunctionName -like '*Get-StaleUsers*'}).rundate | Select-Object -Last 1)
 - Make sure the Variable is set to 'lastscriptname' and that it searches that script name.

2. Create a checkbox for the script. - ex. $cb_DisableStaleUsers = Add-CheckBox 700 30 20 100 "Disable-StaleUsers - Monthly - Last Ran on $LastDisableStaleUsers" ; $MainForm.Controls.Add($cb_DisableStaleUsers)
 - Make sure the y position is set to the last one + 40.              (It goes xsize,ysize,xpos,ypos)

3. Add the If statements for what to show in the text upon startup (example below)
If($LastDisableStaleUsers -eq $null){$cb_DisableStaleUsers.Text = "Disable-StaleUsers - Monthly - No Log Found" ; $cb_DisableStaleUsers.BackColor = 'LightPink' ; $cb_DisableStaleUsers.Font = $FontBold ; $cb_DisableStaleUsers.Checked = $true} 
ElseIf((New-TimeSpan -End (get-date) -Start $LastDisableStaleUsers).days -gt 30){$cb_DisableStaleUsers.BackColor = 'LightPink' ; $cb_DisableStaleUsers.Checked = $true ; $cb_DisableStaleUsers.Font = $FontBold}
- The only thing you should need to change is the variables and the script names.

4. Add the checkbox variable to the $AllCheckboxes array. 
- The checkbox variable is the one you just created ex. $cb_GetStaleUsers

5. Add the if statement to run the script if the checkbox is checked (Example Below)
    If ($cb_GetStaleUsers.Checked -eq $true)   {
        Try{start-process powershell.exe -argumentlist "-NoExit Get-StaleUsers"
            $cb_GetStaleUsers.BackColor = 'LightGreen' ; $cb_GetStaleUsers.Font = $Font ; $cb_GetStaleUsers.Text = "Get-StaleUsers - Monthly - Last Ran on $(Get-Date -f 'MM/d/yyyy hh:mm:ss')"}
        Catch{$cb_GetStaleUsers.Back = 'Red' ; $cb_GetStaleUsers.Font = $FontBold ; $cb_GetStaleUsers.Text = "Get-StaleUsers : ERROR PLEASE SEE POWERSHELL TEAM"
            Add-Content -Path $logpath -Value "`ERROR on Get_StaleUsers,$(Get-Date -f 'MM/d/yyyy hh:mm:ss'),$(whoami),$($Error[0])"}
        }
- Again, you should only need to change the variables and the script names. Everything else should be the same.

6. Make sure to test functionality of what you added before assumming completion. If you have issues, please contact Mike Sprous or Jerry Firebaugh for help.

#>
#-------------------------------------------------------------#

#region VARIABLES

## Primary Variables ##
$Add = $null
$Network = 'DC=newton, DC=pentagon, DC=mil'
$logpath = "\\newton\admin\SysAdmin Powershell Module\Scripts\BattleRythm\Logs\Log_BattleRhythm.csv"
$hour = get-date -format HH
# set [int]$hour to a number between 1 and 23 to test functions that are enabled or disabled during certain shifts
# comment out the test value to make the script active
#[int]$hour = "23"



## Font Configurations ##
$FontBold = [System.Drawing.Font]::new("Cosmic Sans", 12, [System.Drawing.FontStyle]::Bold)
$Font = [System.Drawing.Font]::New("Cosmic Sans", 12)



## WinForm Shortcut Variables ##
$Location = 'System.Drawing.Point'
$Size = 'System.Drawing.Size'
$Form = 'System.Windows.Forms.Form'
$Button = 'System.Windows.Forms.Button'
$CheckBox = 'System.Windows.Forms.CheckBox'
$Label = 'System.Windows.Forms.Label'
$PictureBox = 'System.Windows.Forms.PictureBox'
$ListView = 'System.Windows.Forms.ListView'
$ListViewItem = 'System.Windows.Forms.ListViewItem'


## Gathering the Last Run Dates ##
$lastrun = @(import-csv $logpath)
$LastCheckDriveSpace = (($lastrun | Where-Object {$_.FunctionName -like '*Check-DriveSpace*'}).rundate | Select-Object -Last 1)
$LastDisableStaleAdmins = (($lastrun | Where-Object {$_.FunctionName -like '*Disable-StaleAdmins*'}).rundate | Select-Object -Last 1)
$LastDisableStaleUsers = (($lastrun | Where-Object {$_.FunctionName -like '*Disable-StaleUsers*'}).rundate | Select-Object -Last 1)
$LastGetDomainPrinters = (($lastrun | Where-Object {$_.FunctionName -like '*Get-DomainPrinters*'}).rundate | Select-Object -Last 1)
$LastGetExpiredCerts = (($lastrun | Where-Object {$_.FunctionName -like '*Get-ExpiredCerts*'}).rundate | Select-Object -Last 1)
$LastGetStaleAdmins = (($lastrun | Where-Object {$_.FunctionName -like '*Get-StaleAdmins*'}).rundate | Select-Object -Last 1)
$LastGetStaleUsers = (($lastrun | Where-Object {$_.FunctionName -like '*Get-StaleUsers*'}).rundate | Select-Object -Last 1)
$LastGetUptimeReport = (($lastrun | Where-Object {$_.FunctionName -like '*Get-UptimeReport*'}).rundate | Select-Object -Last 1)
$LastTestExchangeServerHealth = (($lastrun | Where-Object {$_.FunctionName -like '*Test-ExchangeServerHealth*'}).rundate | Select-Object -Last 1)
$LastRunHashRefresh = (($lastrun | Where-Object {$_.FunctionName -like '*Run-HashRefresh*'}).RunDate | Select-Object -Last 1)
$LastCheckSvcLastPassSet = (($lastrun | Where-Object {$_.FunctionName -like '*Check-SvcLastPassSet*'}).RunDate | Select-Object -Last 1)

#endregion


#region Commands for Button / Label Creation

Function Add-Button{
Param([int]$xsize,[int]$ysize,[int]$xpos,[int]$ypos,[string]$text)
$bname = New-Object $Button
$bname.Size = New-Object $Size($xsize,$ysize)
$bname.Location = New-Object $Location($xpos,$ypos)
$bname.Text = $text
$bname.Font = $Font 
$bname
}

Function Add-Label{
Param([int]$xsize,[int]$ysize,[int]$xpos,[int]$ypos,[string]$text,[string]$color,[switch]$Fixed3D)
$lname = New-Object $Label
$lname.Size = New-Object $Size($xsize,$ysize)
$lname.Location = New-Object $Location($xpos,$ypos)
If(!($Fixed3D)){$lname.BorderStyle = 'FixedSingle'}
$lname.Font = $Font_MainForm
$lname.TextAlign = 'MiddleCenter'
$lname.BackColor = 'Gray'
$lname.ForeColor = $color
$lname.Text = $text
$lname
}

Function Add-CheckBox{
Param([int]$xsize,[int]$ysize,[int]$xpos,[int]$ypos,[string]$text)
$cbname = New-Object $CheckBox
$cbname.Size = New-Object $Size($xsize,$ysize)
$cbname.Location = New-Object $Location($xpos,$ypos)
$cbname.Text = $text
$cbname.Font = $Font
$cbname.BackColor = 'LightGreen'
#$cbname.TextAlign = 'BottomLeft'
$cbname
}

Function Add-ListView{
Param([int]$xsize,[int]$ysize,[int]$xpos,[int]$ypos)
$lvname = New-Object $ListView
$lvname.Size = New-Object $Size($xsize,$ysize) 
$lvname.Location = New-Object $Location($xpos,$ypos)
$lvname.Font = $Font
$lvname
}

Function Add-Background{
Param([int]$xsize,[int]$ysize,[int]$xpos,[int]$ypos,[string]$color)
$bkgname = New-Object $PictureBox
$bkgname.Size = New-Object $Size($xsize,$ysize)
$bkgname.Location = New-Object $Location($xpos,$ypos)
$bkgname.BackColor = $color
$bkgname.BorderStyle = 'Fixed3D'
$bkgname.SendToBack()
$bkgname
}

#endregion




#region Functions

Function Select-Deselect{
Param([switch]$Select,[switch]$Deselect)
    
    If($Select){ForEach($_ in $AllCheckBoxes){If($_.Enabled -eq $true){If($_.Checked -eq $False){$_.Checked = $true}}}}
    If($Deselect){ForEach($_ in $AllCheckBoxes){If($_.Checked -eq $True){$_.Checked = $False}}}
}

Function RunChecked {
    If ($cb_CheckDriveSpace.Checked -eq $true) {
        Try{start-process powershell.exe -argumentlist "-NoExit Check-DriveSpace" 
            $cb_CheckDriveSpace.BackColor = 'LightGreen' ; $cb_CheckDriveSpace.Font = $Font ; $cb_CheckDriveSpace.Text = "Check-DriveSpace - Monthly - Last Ran on $(Get-Date -Format 'MM/d/yyyy hh:mm:ss')"}
        Catch{$cb_CheckDriveSpace.BackColor = 'Red' ; $cb_CheckDriveSpace.Font = $FontBold ; $cb_CheckDriveSpace.Text = "Check-DriveSpace : ERROR PLEASE SEE POWERSHELL TEAM"
            Add-Content -Path $logpath -Value "`ERROR on Check_DriveSpace,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami),$($Error[0])"}
        }

    If ($cb_GetStaleAdmins.Checked -eq $true)  {
        Try{start-process powershell.exe -argumentlist "-NoExit Get-StaleAdmins" 
            $cb_GetStaleAdmins.BackColor = 'LightGreen' ; $cb_GetStaleAdmins.Font = $Font ; $cb_GetStaleAdmins.Text = "Get-StaleAdmins - Monthly - Last Ran on $(Get-Date -F 'MM/d/yyyy hh:mm:ss')"}
        Catch{$cb_GetStaleAdmins.BackColor = 'Red' ; $cb_GetStaleAdmins.Font = $FontBold ; $cb_GetStaleAdmins.Text = "Get-StaleAdmins : ERROR PLEASE SEE POWERSHELL TEAM"
            Add-Content -Path $logpath -Value "`ERROR on Get_StaleAdmins,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami),$($Error[0])"}
        }

    If ($cb_GetStaleUsers.Checked -eq $true)   {
        Try{start-process powershell.exe -argumentlist "-NoExit Get-StaleUsers -Days 365" 
            $cb_GetStaleUsers.BackColor = 'LightGreen' ; $cb_GetStaleUsers.Font = $Font ; $cb_GetStaleUsers.Text = "Get-StaleUsers - Monthly - Last Ran on $(Get-Date -f 'MM/d/yyyy hh:mm:ss')"}
        Catch{$cb_GetStaleUsers.Back = 'Red' ; $cb_GetStaleUsers.Font = $FontBold ; $cb_GetStaleUsers.Text = "Get-StaleUsers : ERROR PLEASE SEE POWERSHELL TEAM"
            Add-Content -Path $logpath -Value "`ERROR on Get_StaleUsers,$(Get-Date -f 'MM/d/yyyy hh:mm:ss'),$(whoami),$($Error[0])"}
        }

    If ($cb_DisableStaleAdmins.Checked -eq $true) { If($cb_DisableStaleAdmins.Enabled -eq $False){}Else{
        Try{start-process powershell.exe -argumentlist "-NoExit Disable-StaleAdmins"  
            $cb_DisableStaleAdmins.BackColor = 'LightGreen' ; $cb_DisableStaleAdmins.Font = $Font ; $cb_DisableStaleAdmins.Text = "Disable-StaleAdmins - Monthly - Last Ran on $(get-date -Format 'MM/d/yyyy hh:mm:ss')"}
        Catch{$cb_DisableStaleAdmins.BackColor = 'Red' ; $cb_DisableStaleAdmins.Font = $FontBold ; $cb_DisableStaleAdmins.Text = "Disable-StaleAdmins : ERROR PLEASE SEE POWERSHELL TEAM"
            Add-Content -Path $logpath -Value "`ERROR on Disable_StaleAdmins,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami),$($Error[0])"}
            }
        }

    If ($cb_DisableStaleUsers.Checked -eq $true) {If($cb_DisableStaleUsers.Enabled -eq $false){}Else{
        Try{start-process powershell.exe -argumentlist "-NoExit Disable-StaleUsers" 
            $cb_DisableStaleUsers.BackColor = 'LightGreen' ; $cb_DisableStaleUsers.Font = $Font ; $cb_DisableStaleUsers.Text = "Disable-StaleUsers - Monthly - Last Ran on $(Get-Date -Format 'MM/d/yyyy hh:mm:ss')"}
        Catch{$cb_DisableStaleUsers.BackColor = 'Red' ; $cb_DisableStaleUsers.Font = $FontBold ; $cb_DisableStaleUsers.Text = "Disable-StaleUsers : ERROR PLEASE SEE POWERSHELL TEAM"
            Add-Content -Path $logpath -Value "`ERROR on Disable_StaleUsers,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami),$($Error[0])"}
            }
        }

    If ($cb_GetDomainPrinters.Checked -eq $true) {
        Try{start-process powershell.exe -argumentlist "-NoExit Get-DomainPrinters" 
            $cb_GetDomainPrinters.BackColor = 'LightGreen' ; $cb_GetDomainPrinters.Font = $Font ; $cb_GetDomainPrinters.Text = "Get-DomainPrinters - Monthly - Last Ran on $(Get-Date -Format 'MM/d/yyyy hh:mm:ss')"} 
        Catch{$cb_GetDomainPrinters.BackColor = 'Red' ; $cb_GetDomainPrinters.Font = $FontBold ; $cb_GetDomainPrinters.Text = "Get-DomainPrinters : ERROR PLEASE SEE POWERSHELL TEAM"
            Add-Content -Path $logpath -Value "`ERROR on Get_DomainPrinters,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami),$($Error[0])"}
        }

    If ($cb_GetExpiredCerts.Checked -eq $true) {
        Try{start-process powershell.exe -argumentlist "-NoExit Get-ExpiredCerts"   
            $cb_GetExpiredCerts.BackColor = 'LightGreen' ; $cb_GetExpiredCerts.Font = $Font ; $cb_GetExpiredCerts.Text = "Get-ExpiredCerts - Weekly - Last Ran on $(Get-Date -Format 'MM/d/yyyy hh:mm:ss')"}
        Catch{$cb_GetExpiredCerts.BackColor = 'Red' ; $cb_GetExpiredCerts.Font = $FontBold ; $cb_GetExpiredCerts.Text = "Get-ExpiredCerts : ERROR PLEASE SEE POWERSHELL TEAM"
            Add-Content -Path $logpath -Value "`ERROR on Get_ExpiredCerts,$(Get-Date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami),$($Error[0])"}
        }

    If ($cb_GetUptimeReport.Checked -eq $true) {
        Try{start-process powershell.exe -argumentlist "-NoExit Get-UptimeReport -allservers" 
            $cb_GetUptimeReport.BackColor = 'LightGreen' ; $cb_GetUptimeReport.Font = $Font ; $cb_GetUptimeReport.Text = "Get-UptimeReport - Monthly - Last Ran on $(Get-Date -Format 'MM/d/yyyy hh:mm:ss')"}
        Catch{$cb_GetUptimeReport.BackColor = 'Red' ; $cb_GetUptimeReport.Font = $FontBold ; $cb_GetUptimeReport.Text = "Get-UptimeReport : ERROR PLEASE SEE POWERSHELL TEAM"
            Add-Content -Path $logpath -Value "`ERROR on Get_UptimeReport,$(Get-Date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami),$($Error[0])"}
        }

    If ($cb_TestExchangeServerHealth.Checked -eq $true) {
        Try{start-process powershell.exe -argumentlist "-NoExit Test-ExchangeServerHealth -sendemail" 
            $cb_TestExchangeServerHealth.BackColor = 'LightGreen' ; $cb_TestExchangeServerHealth.Font = $Font ; $cb_TestExchangeServerHealth.Text = "Test-ExchangeServerHealth - Each Shift - Last Ran on $(Get-Date -Format 'MM/d/yyyy hh:mm:ss')"}
        Catch{$cb_TestExchangeServerHealth.BackColor = 'Red' ; $cb_TestExchangeServerHealth.Font = $FontBold ; $cb_TestExchangeServerHealth.Text = "Test-ExchangeServerHealth : ERROR PLEASE SEE POWERSHELL TEAM"
            Add-Content -Path $logpath -Value "`ERROR on Test_ExchangeServerHealth,$(Get-Date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami),$($Error[0])"}
        }

    If($cb_RunHashRefresh.Checked -eq $true) {
        $Confirm = [System.Windows.MessageBox]::Show('You are about to run the hash refresh. Please confirm this action before continuing','Confirmation','YesNo')
        If($Confirm -eq 'Yes'){
        Try{Start-Process powershell.exe -ArgumentList "-NoExit Run-HashRefresh"
            $cb_RunHashRefresh.BackColor = 'LightGreen' ; $cb_RunHashRefresh.Font = $Font ; $cb_RunHashRefresh.Text = "Run-HashRefresh - Monthly Mids Only - Last Ran on $(Get-Date -Format 'MM/d/yyyy hh:mm:ss')"}
        Catch{$cb_RunHashRefresh.BackColor = 'Red' ; $cb_RunHashRefresh.Font = $FontBold ; $cb_RunHashRefresh.Text = "Run-HashRefresh : ERROR PLEASE SEE POWERSHELL TEAM"
            Add-Content -Path $logpath -Value "`ERROR on Run_HashRefresh,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami),$($Error[0])"}
            }Else{}
        }

    If($cb_CheckSvcLastPassSet.Checked -eq $true) {
        Try{Start-Process powershell.exe -ArgumentList "-NoExit Check-SvcLastPassSet"
            $cb_CheckSvcLastPassSet.BackColor = 'LightGreen' ; $cb_CheckSvcLastPassSet.Font = $Font ; $cb_CheckSvcLastPassSet.Text = "Check-SvcLastPassSet - Annually - Last Ran on $(Get-Date -Format 'MM/d/yyyy hh:mm:ss')"}
        Catch{$cb_CheckSvcLastPassSet.BackColor = 'Red' ; $cb_CheckSvcLastPassSet.Font = $FontBold ; $cb_CheckSvcLastPassSet.Text = "Check-SvcLastPassSet : ERROR PLEASE SEE POWERSHELL TEAM"
            Add-Content -Path $logpath -Value "`ERROR on Check_SvcLastPassSet,$(Get-Date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami),$($Error[0])"}
        }

}

Function Review-Accounts {
Param([switch]$Admins,[switch]$Users)

    If($Admins){
        $folder = "\\newton\admin\Scripts\Monthly Scripts\Stale Admin Accounts\Logs"
        $FolderList = Get-ChildItem -Path $Folder | Sort LastWriteTime -Descending ; $ReportFolder = $FolderList[0].FullName
        $File = Get-ChildItem -Path $ReportFolder | Sort LastWriteTime -Descending | Where Name -like "*.csv*" ; $CSV = $File[0].FullName
        Invoke-Item -Path $ReportFolder
        Invoke-Item -Path $CSV

        # Unlock the Checkbox
        $cb_DisableStaleAdmins.Enabled = $true
    }

    If($Users){
        $folder = "\\newton\admin\Scripts\Monthly Scripts\Stale User Accounts\Logs"
        $FolderList = Get-ChildItem -Path $Folder | Sort LastWriteTime -Descending ; $ReportFolder = $FolderList[0].FullName
        $File = Get-ChildItem -Path $ReportFolder | Sort LastWriteTime -Descending | Where Name -like "*.csv*" ; $CSV = $File[0].FullName
        Invoke-Item -Path $ReportFolder
        Invoke-Item -Path $CSV
        
        # Unlock the Checkbox
        $cb_DisableStaleUsers.Enabled = $true        
    }
}

Function Get-LatestSVC {
$Folder = "\\newton\admin\SysAdmin Powershell Module\Scripts\Check-SvcLastPassSet\Logs"
$SvcAccts = Get-ADUser -Filter {(SamAccountName -like "*newton-sysadmin*") -or (SamAccountName -like "*SVC*") -and (Enabled -eq $True)} -Properties PasswordLastSet,whenCreated
$Yr = Get-Date -Format MM-yyyy
$Report = @()
Remove-item -Path $Folder\ExpiredPassword-Svc-$Yr.csv -ErrorAction SilentlyContinue

# - Go through each service account and if there was no date for a password last set, it will be changed to the date the account was created.
ForEach($fp in $SvcAccts){ If($fp.PasswordLastSet -eq $null){ $fp.PasswordLastSet = $s.WhenCreated } }

# - Go through each service account and add each one that hasn't been changed in 350 days or more to the list of bad ones.
ForEach($fp in $SvcAccts){
    If($fp.PasswordLastSet -lt (Get-Date).adddays(-350)){ 
        $Report += [pscustomobject]@{ SamAccountName = $fp.SamAccountName ; WhenCreated = $fp.whenCreated ; PasswordLastSet = $fp.PasswordLastSet }
        $c = $fp.whenCreated ; $l = $fp.PasswordLastSet
        $Add = New-Object $ListViewItem($fp.SamAccountName) ; $Created = $Add.Subitems.Add("$c") ; $LastChange = $Add.Subitems.Add("$l")
        $lv_svcaccounts.Items.AddRange($Add)
        }
    }
$Report | Select SamAccountName,WhenCreated,PasswordLastSet | Export-Csv -Path $Folder\ExpiredPassword-Svc-$Yr.csv -NoType -Encoding UTF8 -Append
$Report | Select SamAccountName,WhenCreated,PasswordLastSet | clip

# - Checks if there were service accounts found and display the result.
#If(!(Test-Path -Path $Folder\ExpiredPassword-Svc-$Yr.csv -ErrorA Si)){}Else{Invoke-Item -Path $Folder\ExpiredPassword-Svc-$Yr.csv}

#$logpath = "\\newton\admin\SysAdmin Powershell Module\Scripts\BattleRythm\Logs\Log_BattleRhythm.csv"
#Add-Content -Path $logpath -Value "`Check-SvcLastPassSet,$(get-date -Format 'MM/d/yyyy hh:mm:ss'),$(whoami)"
}

#endregion




#region Form Contrustion

<# Main Form #>
$MainForm = New-Object $Form ; $MainForm.Size = New-Object $Size(740,700) ; $MainForm.Text = 'Battle Rhythm Scheduled Task Executor' # - Adds the Main Form 
$MainForm.FormBorderStyle = 'Fixed3D' ; 

<# Checkboxes #>
# - Check-DriveSpace Checkbox
$cb_CheckDriveSpace = Add-CheckBox 700 30 10 20 "Check-DriveSpace - Weekly - Last Ran on $LastCheckDriveSpace" ; $MainForm.Controls.Add($cb_CheckDriveSpace) 
If($LastCheckDriveSpace -eq $null){$cb_CheckDriveSpace.Text = "Check-DriveSpace - Monthly - No Log Found" ; $cb_CheckDriveSpace.BackColor = 'LightPink' ; $cb_CheckDriveSpace.Font = $FontBold ; $cb_CheckDriveSpace.Checked = $true}
ElseIf((New-TimeSpan -End (get-date) -Start $LastCheckDriveSpace).days -gt 7){$cb_CheckDriveSpace.BackColor = 'LightPink' ; $cb_CheckDriveSpace.Checked = $true ; $cb_CheckDriveSpace.Font = $FontBold}

# - Disable-StaleAdmins Checkbox
$cb_DisableStaleAdmins = Add-CheckBox 700 30 10 60 "Disable-StaleAdmins - Weekly - Last Ran on $LastDisableStaleAdmins" ; $MainForm.Controls.Add($cb_DisableStaleAdmins) 
If($LastDisableStaleAdmins -eq $null){$cb_DisableStaleAdmins.Text = "Disable-StaleAdmins - Monthly - No Log Found" ; $cb_DisableStaleAdmins.BackColor = 'LightPink' ; $cb_DisableStaleAdmins.Font = $FontBold ; $cb_DisableStaleAdmins.Checked = $true}
ElseIf((New-TimeSpan -End (get-date) -Start $LastDisableStaleAdmins).days -gt 30){$cb_DisableStaleAdmins.BackColor = 'LightPink' ; $cb_DisableStaleAdmins.Checked = $true ; $cb_DisableStaleAdmins.Font = $FontBold}

# - Disable-StaleUsers Checkbox
$cb_DisableStaleUsers = Add-CheckBox 700 30 10 100 "Disable-StaleUsers - Weekly - Last Ran on $LastDisableStaleUsers" ; $MainForm.Controls.Add($cb_DisableStaleUsers)
If($LastDisableStaleUsers -eq $null){$cb_DisableStaleUsers.Text = "Disable-StaleUsers - Monthly - No Log Found" ; $cb_DisableStaleUsers.BackColor = 'LightPink' ; $cb_DisableStaleUsers.Font = $FontBold ; $cb_DisableStaleUsers.Checked = $true} 
ElseIf((New-TimeSpan -End (get-date) -Start $LastDisableStaleUsers).days -gt 30){$cb_DisableStaleUsers.BackColor = 'LightPink' ; $cb_DisableStaleUsers.Checked = $true ; $cb_DisableStaleUsers.Font = $FontBold}

# - Get-DomainPrinters Checkbox
$cb_GetDomainPrinters = Add-CheckBox 700 30 10 140 "Get-DomainPrinters - Monthly - Last Ran on $LastGetDomainPrinters" ; $MainForm.Controls.Add($cb_GetDomainPrinters) 
If($LastGetDomainPrinters -eq $null){$cb_GetDomainPrinters.Text = "Get-DomainPrinters - Monthly - No Log Found" ; $cb_GetDomainPrinters.BackColor = 'LightPink' ; $cb_GetDomainPrinters.Font = $FontBold ; $cb_GetDomainPrinters.Checked = $true}
ElseIf((New-TimeSpan -End (get-date) -Start $LastGetDomainPrinters).days -gt 30){$cb_GetDomainPrinters.BackColor = 'LightPink' ; $cb_GetDomainPrinters.Checked = $true ; $cb_GetDomainPrinters.Font = $FontBold}

# - Get-ExpiredCerts Checkbox
$cb_GetExpiredCerts = Add-CheckBox 700 30 10 180 "Get-ExpiredCerts - Weekly - Last Ran on $LastGetExpiredCerts" ; $MainForm.Controls.Add($cb_GetExpiredCerts) 
If($LastGetExpiredCerts -eq $null){$cb_GetExpiredCerts.Text = "Get-ExpiredCerts - Weekly - No Log Found" ; $cb_GetExpiredCerts.BackColor = 'LightPink' ; $cb_GetExpiredCerts.Font = $FontBold ; $cb_GetExpiredCerts.Checked = $true}
ElseIf((New-TimeSpan -End (get-date) -Start $LastGetExpiredCerts).days -gt 7){$cb_GetExpiredCerts.BackColor = 'LightPink' ; $cb_GetExpiredCerts.Checked = $true ; $cb_GetExpiredCerts.Font = $FontBold}

# - Get-StaleAdmins Checkbox
$cb_GetStaleAdmins = Add-CheckBox 700 30 10 220 "Get-StaleAdmins - Weekly - Last Ran on $LastGetStaleAdmins" ; $MainForm.Controls.Add($cb_GetStaleAdmins)
If($LastGetStaleAdmins -eq $null){$cb_GetStaleAdmins.Text = "Get-StaleAdmins - Monthly - No Log Found" ; $cb_GetStaleAdmins.BackColor = 'LightPink' ; $cb_GetStaleAdmins.Font = $FontBold ; $cb_GetStaleAdmins.Checked = $true} 
ElseIf((New-TimeSpan -End (get-date) -Start $LastGetStaleAdmins).days -gt 30){$cb_GetStaleAdmins.BackColor = 'LightPink' ; $cb_GetStaleAdmins.Checked = $true ; $cb_GetStaleAdmins.Font = $FontBold}

# - Get-StaleUsers Checkbox
$cb_GetStaleUsers = Add-CheckBox 700 30 10 260 "Get-StaleUsers - Weekly - Last Ran on $LastGetStaleUsers" ; $MainForm.Controls.Add($cb_GetStaleUsers)
If($LastGetStaleUsers -eq $null){$cb_GetStaleUsers.Text = "Get-StaleUsers - Monthly - No Log Found" ; $cb_GetStaleUsers.BackColor = 'LightPink' ; $cb_GetStaleUsers.Font = $FontBold ; $cb_GetStaleUsers.Checked = $true} 
ElseIf((New-TimeSpan -End (get-date) -Start $LastGetStaleUsers).days -gt 30){$cb_GetStaleUsers.BackColor = 'LightPink' ; $cb_GetStaleUsers.Checked = $true ; $cb_GetStaleUsers.Font = $FontBold}

# - Get-UptimeReport Checkbox
$cb_GetUptimeReport = Add-CheckBox 700 30 10 300 "Get-UptimeReport - Monthly - Last Ran on $LastGetUptimeReport" ; $MainForm.Controls.Add($cb_GetUptimeReport)
If($LastGetUptimeReport -eq $null){$cb_GetUptimeReport.Text = "Get-UptimeReport - Monthly - No Log Found" ; $cb_GetUptimeReport.BackColor = 'LightPink' ; $cb_GetUptimeReport.Font = $FontBold ; $cb_GetUptimeReport.Checked = $true}
ElseIf((New-TimeSpan -End (get-date) -Start $LastGetUptimeReport).days -gt 30){$cb_GetUptimeReport.BackColor = 'LightPink' ; $cb_GetUptimeReport.Checked = $true ; $cb_GetUptimeReport.Font = $FontBold}

# - Test-ExchangeServerHealth Checkbox
$cb_TestExchangeServerHealth = Add-CheckBox 700 30 10 340 "Test-ExchangeServerHealth - Every 4 Hours - Last Ran on $LastTestExchangeServerHealth" ; $MainForm.Controls.Add($cb_TestExchangeServerHealth)
If($LastTestExchangeServerHealth -eq $null){$cb_TestExchangeServerHealth.Text = "Test-ExchangeServerHealth - Each Shift - No Log Found" ; $cb_TestExchangeServerHealth.BackColor = 'LightPink' ; $cb_TestExchangeServerHealth.Font = $FontBold ; $cb_TestExchangeServerHealth.Checked = $true}
ElseIf((New-TimeSpan -End (get-date) -Start $LastTestExchangeServerHealth).Totalhours -gt 6){$cb_TestExchangeServerHealth.BackColor = 'LightPink' ; $cb_TestExchangeServerHealth.Checked = $true ; $cb_TestExchangeServerHealth.Font = $FontBold}

# - Check-SvcLastPassSet
$cb_CheckSvcLastPassSet = Add-CheckBox 700 30 10 380 "Check-SvcLastPassSet - Annually - Last Ran on $LastCheckSvcLastPassSet" ; $MainForm.Controls.Add($cb_CheckSvcLastPassSet)
If($LastCheckSvcLastPassSet -eq $null){$cb_CheckSvcLastPassSet.Text = "Check-SvcLastPassSet - No Log Found" ; $cb_CheckSvcLastPassSet.BackColor = 'LightPink' ; $cb_CheckSvcLastPassSet.Font = $FontBold ; $cb_CheckSvcLastPassSet.Checked = $true}
ElseIf((New-TimeSpan -End (Get-Date) -Start $LastCheckSvcLastPassSet).Days -gt 364){$cb_CheckSvcLastPassSet.BackColor = 'LightPink' ; $cb_CheckSvcLastPassSet.Checked = $True ; $cb_CheckSvcLastPassSet.Font = $FontBold}



# Array for All Checkboxes
$AllCheckBoxes = ($cb_CheckDriveSpace,$cb_DisableStaleAdmins,$cb_DisableStaleUsers,$cb_GetDomainPrinters,$cb_GetExpiredCerts,$cb_GetStaleAdmins,$cb_GetStaleUsers,$cb_GetUptimeReport,$cb_TestExchangeServerHealth,
                  $cb_CheckSvcLastPassSet)

<# Buttons #>
# Check All Button
$btn_CheckAll = Add-Button 110 50 20 590 "Select All" ; $MainForm.Controls.Add($btn_CheckAll) 
$btn_CheckAll.Add_Click({Select-Deselect -Select})

# Uncheck all Button
$btn_UnCheckAll = Add-Button 110 50 140 590 "Deselect All" ; $MainForm.Controls.Add($btn_UnCheckAll) 
$btn_UnCheckAll.Add_Click({Select-Deselect -Deselect})

# Run Select Button
$btn_RunChecked = Add-Button 120 50 340 590 "Run Selected" ; $MainForm.Controls.Add($btn_RunChecked) 
$btn_RunChecked.Add_Click({RunChecked})

# Help Button
$btn_Help = Add-Button 100 50 480 590 "Help" ; $MainForm.Controls.Add($btn_Help)
$btn_Help.Add_Click({start-process powershell.exe -NoNewWindow -ArgumentList "Help-SysAdmin"})

# Exit Button
$btn_Exit = Add-Button 100 50 600 590 "Exit" ; $MainForm.Controls.Add($btn_Exit) 
$btn_Exit.Add_Click({$MainForm.Dispose()})

# Review Disabled Admins
$btn_Review_Admins = Add-Button 150 25 557 62 "Review Required" ; $MainForm.Controls.Add($btn_Review_Admins) ; $btn_Review_Admins.BringToFront()
$btn_Review_Admins.Font = $FontBold ; $btn_Review_Admins.Add_Click({Review-Accounts -Admins})

# Review Disabled Users
$btn_Review_Users = Add-Button 150 25 557 102 "Review Required" ; $MainForm.Controls.Add($btn_Review_Users) ; $btn_Review_Users.BringToFront()
$btn_Review_Users.Font = $FontBold ; $btn_Review_Users.Add_Click({Review-Accounts -Users})


<# List Boxes #>
$lv_svcaccounts = Add-ListView 700 100 10 465 ; $lv_svcaccounts.View = 'Details' ; $MainForm.Controls.Add($lv_svcaccounts)
$column1 = $lv_svcaccounts.Columns.Add('Account Name') ; $column1.Width = 200
$column2 = $lv_svcaccounts.Columns.Add('Creation Date') ; $column2.Width = 245
$column3 = $lv_svcaccounts.Columns.Add('Last Password Changed') ; $column3.Width = 230 
Get-LatestSVC


<# Backgrounds #>
# Menu Options Background
$bkg_MenuOptions = Add-Background 700 70 10 580 LightGray ; $MainForm.Controls.Add($bkg_MenuOptions)

#endregion








# Decide whether Disabled Admins/Users needs to be reviewed or not. 
If($cb_DisableStaleAdmins.BackColor -eq 'LightPink'){$btn_Review_Admins.Visible = $true ; $cb_DisableStaleAdmins.Enabled = $false ; $cb_DisableStaleAdmins.Checked = $false}
    Else{$btn_Review_Admins.Visible = $false}
If($cb_DisableStaleUsers.BackColor -eq 'LightPink'){$btn_Review_Users.Visible = $true ; $cb_DisableStaleUsers.Enabled = $false ; $cb_DisableStaleUsers.Checked = $false}
    Else{$btn_Review_Users.Visible = $false}

If($Add -ge 1){$cb_CheckSvcLastPassSet.Text = "Check-SvcLastPassSet - No Violations Found - Last Ran on $LastCheckSvcLastPassSet" ; $cb_CheckSvcLastPassSet.BackColor = 'LightGreen' ; $cb_CheckSvcLastPassSet.Font = $Font}
Else{$cb_CheckSvcLastPassSet.Text = "Check-SvcLastPassSet - Violations Found - Last Ran on $LastCheckSvcLastPassSet" ; $cb_CheckSvcLastPassSet.BackColor = 'LightPink' ; $cb_CheckSvcLastPassSet.Font = $FontBold ; $cb_CheckSvcLastPassSet.Checked = $True}


#LAST LINE HERE
[System.Windows.Forms.Application]::Run($MainForm)
}

Function Update-SysAdmin {
[cmdletbinding()]
Param(
    [Parameter (Mandatory = $False)][switch]$Single,
    [Parameter (Mandatory = $false)][switch]$Refresh
)

$Module = Get-Item -Path '\\newton\admin\SysAdmin Powershell Module\SysAdmin.psm1'
$Manifest = Get-Item -Path '\\newton\admin\SysAdmin Powershell Module\SysAdmin.psd1'
#$Scripts = Get-Item -Path '\\newton\admin\SysAdmin Powershell Module\Scripts'
$Count = 0
$Comp = ""
$Path = "\\newton\admin\SysAdmin Powershell Module\Scripts\Update-SysAdmin\SysAdminComputers.csv"
$SysAdminComps = Import-Csv -Path "\\newton\admin\SysAdmin Powershell Module\Scripts\Update-SysAdmin\SysAdminComputers.csv"
$Options = @('y','n','Y','N')


If($Single){
    $Computers = Read-Host "Please Enter a Computer Name"
    Do {
        $Choice = Read-Host "Do you want to add this computer to the SysAdmin Computer List? (Enter Y or N)" ; $answered = $False
        If($Options -notcontains $Choice){Write-Host "Please Enter Y or N" -f Yellow -b Black}
        Else{
            Switch($Choice){
            'y' {$report = [pscustomobject]@{Computer = $Computers} ; $report | select Computer | Export-csv -Path $path -NoTypeInformation -encoding UTF8 -Append ; $answered = $true}
            'n' {$answered = $true}
            }
            }
        }Until($answered -eq $true)
    } 

Else{
    $Computers = $SysAdminComps
    }

$Total = $Computers.Count

ForEach($Compu in $Computers){ 
If($Single){$Comp = $Compu}Else{$Comp = $compu.Computer}
If($Comp -eq $env:COMPUTERNAME){$Location = "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\SysAdmin"}
Else{$Location = "\\$Comp\C$\Windows\System32\WindowsPowerShell\v1.0\Modules\SysAdmin"}
$Count++
Write-Host "(" -f DarkGray -no ; Write-Host $Count -f Yellow -no ; Write-Host " of " -f DarkGray -no ; Write-Host $Total -f Yellow -no ; Write-Host ") " -f DarkGray -no ; Write-Host $Comp -f Cyan
Write-Host "- " -f DarkGray -no ; Write-Host "Module Directory : " -f Yellow -no 
If(!(Test-Path $Location -ErrorAction Si)){$New = New-Item -Path $Location -ItemType Directory -InformationAction Si ; Write-Host "Created." -f Green}Else{Write-Host "Good." -f Green}

Write-Host "- " -f DarkGray -no ; Write-Host "Updating Module : " -f Yellow -No
Copy-Item $Module -Destination $Location -Force ; Write-Host "Done." -f Green

Write-Host "- " -f DarkGray -no ; Write-Host "Updating Manifest : " -f Yellow -no 
Copy-Item $Manifest -Destination $Location -Force ; Write-Host "Done." -f Green

#Write-Host "- " -f DarkGray -no ; Write-Host "Updating Scripts : " -f Yellow -no
#Copy-Item $Scripts -Destination $Location -Recurse -Force ; Write-Host "Done." -f Green

Write-Host "" ; Write-Host ""


    }
Write-Host "Module Updated." -f Green

cd "C:\Windows\System32"

If($Refresh){cls}
powershell.exe

}

Function Help-SysAdmin {
#Loading Assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

<# Form Variables #>
$Location = 'System.Drawing.Point'
$Size = 'System.Drawing.Size'
$Form = 'System.Windows.Forms.Form'
$PictureBox = 'System.Windows.Forms.PictureBox'
$Button = 'System.Windows.Forms.Button'
$CheckBox = 'System.Windows.Forms.Checkbox'
$ComboBox = 'System.Windows.Forms.Combobox'
$ListBox = 'System.Windows.Forms.ListBox'
$CheckedListBox = 'System.Windows.Forms.CheckedListBox'
$RadioButton = 'System.Windows.Forms.RadioButton'
$TextBox = 'System.Windows.Forms.TextBox'
$RichTextBox = 'System.Windows.Forms.RichTextBox'
$Label = 'System.Windows.Forms.Label'
$ProgressBar = 'System.Windows.Forms.ProgressBar'
$Timer = 'System.Windows.Forms.Timer'
$PrimaryFontStyle = [System.Drawing.Font]::new("Cosmic Sans", 12, [System.Drawing.FontStyle]::Bold)
$SecondaryFontStyle = [System.Drawing.Font]::new("Cosmic Sans", 12)
$networkfont = [System.Drawing.Font]::new("Cosmic Sans", 14, [System.Drawing.FontStyle]::Bold)

# Changeable Variables
$ActiveScripts = Get-Content -Path '\\n020fp05\admin\SysAdmin Powershell Module\Active-Scripts.txt'
$fpnfo_GetDomainPrinters = "Get-DomainPrinters is designed to gather information about all the currently connected printers on the domain. The script will output in the powershell console the current printer and try to establish a connection to it. 
Once complete, a report will be placed inside a folder, which can be opened by using the open report folder below.

The information gathered is listed below;
Printer Server - The server that the printer is connected to.
Printer Name - The hostname attached to the printer.
Port - The port assigned to the printer. This will show either green or red depending on the ping results.

Available additional parameters;
-OpenReport : This opens the generated report automatically after completion of the script."

$fpnfo_GetExpiredCerts = "Get-ExpiredCerts goes through the certificate store of every Windows Server on the domain and gathers information on their certificates.
If a certificate is expired, it needs to be replaced / renewed. This can be done by any admin on Lowside, but only by Tier 3 Admins on Snet and TSnet.

Information Gathered listed below;
Certificate Name - The 'subject' or name of the certificate.
Issuer - The CA (Certificate Authority) that issued the certificate.
Expiration Date - The date that the certificate is no longer valid.
Status - Whether the certificate is expired, about to expire, or is still in good standing.

"

$fpnfo_GetStaleUsers = "Get-StaleUSers is designed to gather all the users who have not logged into their account past (x) days. The script can be run by default for 365 days, or be ran to search for one of 5 different searches (30, 60, 90, 180, 365) days.
After the users are gathered, a report is generated in the report folder on the adminshare. It is recommended to run Disable-StaleUsers after this to ensure the accounts found are correctly disabled. 

Information gathered listed below;
Name of User - The name of the individual associated to the account.
Creation Date - When the accoutn was created.
Last Logon - Last time the user logged into the account.
Days since last logon - Days calculated to when the user last logged on.

Additional Parameters;
-OpenReport : Opens the report file immediately after completion of the script.
-Days : Allows the admin to set the criteria of what days to search for in regards to an account being active or not. Options are 30,60,90,180, and 365 days.
"

$fpnfo_GetStaleAdmins = "Get-StaleAdmins is designed to gather all the admin accounts that have not been logged into within the past (x) days. The script will be run by default for 90 days.
After the accounts are gathered, a report is generated in the report folder on the adminshare. It is recommended to run Disable-StaleAdmins after this to ensure the accounts found are correctly disabled.  

Information gathered listed below;
Name of User - The name of the individual associated to the account.
Creation Date - When the accoutn was created.
Last Logon - Last time the account was logged into the account.
Days since last logon - Days calculated to when the account was last logged on.

Additional Parameters;
-OpenReport : Opens the report file immediately after completion of the script."

$fpnfo_DisableStaleUsers = "Disable-StaleUsers allows the administrator to gather the list of stale user accounts and automatically disable them through this script. The script will ask for a ticket number, which it requires. 
As the script processes, it will go through each account and disable the account, then altering the decription to add that it was disabled in reference to the ticket number previously entered.

NOTE : On TSnet, an excel sheet will open upon completion, please be sure to remove the names of the users within their assigned jabber numbers.

Information outputted to console;
Account Username - The username of the account it's currently disabling.
Status - Shows whether the account was disabled or not.
"

$fpnfo_DisableStaleAdmins = "Disable-StaleAdmins allows the administrator to gather the list of stale admin accounts and automatically disable them through this script. The script will ask for a ticket number, which it requires. 
As the script processes, it will go through each account and disable the account, then altering the decription to add that it was disabled in reference to the ticket number previously entered.

NOTE : On TSnet, an excel sheet will open upon completion, please be sure to remove the names of the users within their assigned jabber numbers.

NOTE : On Snet and TSnet, only Tier 3 Admins have permission to run this script.

Information outputted to console;
Account Username - The username of the account it's currently disabling.
Status - Shows whether the account was disabled or not.
"

$fpnfo_ClonePermissions = "Clone-Permissions is designed to help administrators add security groups to user accounts to match other users. This script is to be ran only when authorized as every security group will be added from one account to the other.
It is recommended to make sure the account you are cloning from and cloning to are both enabled and are correct. 

NOTE : Some security groups will require Tier 3 permissions and will be annotated within the powershell console. Please pay attention to this.

Information processed listed below;
Each group will be processed individually, allowing the status of each group to be verified.
"

$fpnfo_ResetBitLockerKey = "Reset-BitLockerKey is designed to gather the bitlocker password of a machine from either AD (Active Directory), or the machine itself.
The script will scan AD for the recovery password first, and if not found or permissions are denied, it will try to gather it from the machine. 
This script is also capable of RESETTING a bitlocker password by Decrypting and Re-Encrypting the machine, this will take some time. Once completed, it will automatically update AD with the new recovery password.

NOTE : If it tries to gather from the machine, the machine will need to be online for it to be successful. So it is recommended a Tier 3 Administrator run this script to avoid complications.

NOTE : When resetting, please be patient and don't cancel the script. Cancelling the script will it's running a reset can break bitlocker and corrupt the drive.

Informaiton Gathered listed below;
Computer Name - The hostname of the computer selected.
Recovery Password - The password used to recovery a machines bitlocker.
"

$fpnfo_checkdrivespace = "Check-DriveSpace is designed to gather information of every drive within the Domain Controllers, File&Print Servers, and Exchange Servers. 

The information gathered is listed below;
Drive Letter - The letter assigned to the drive, if applicable.
Label - The name of the drive.
Size - The total size in GB allocated to the disk.
Free Space - The amount of space in GB currently available.
Percentage Free - The overall percentage of free space available

Once all the information is gathered, an email report will be generated and sent to the SysAdmin distrogroup for everyone to see. This email will list only the drives at 20% or lower free space.

Available additional parameters; 
-ShowOutput : This will show the output within the powershell console for the admin to view as its happening.
"

$fpnfo_TestExchangeServerHealth = "This script is designed to verify the functionality of each Exchange server on the current domain. It will go through every server and verify multiple aspects of exchange, database, and services.
The output of each check is outputted to the powershell console. It's recommended to view this as it's happening to understand better of what is being checked.
Once completed, it will generate an HTML report and send that report to the SysAdmin distrogroup with it's findings. Here you are to verify the findings and troubleshoot any errors. 

NOTE : This is to be ran DAILY, once on EACH SHIFT.
NOTE : Generally, you want to use the parameter '-SendEmail' so it will send the email to the -NOCsysadmins distro group.

Addtional Parameters;
-SendEmail : This will send the report to the -NOCSysadmins distro group
-Server : Specify a specific server to check
-ServerList : Specify a number of specific servers to check
-ReportFile : Open an html page of the report upon completion
-Alertsonly : Script will ignore all good checks and only output checks that need attention
-Log : Created a logfile of the report as a .txt format.
"
$fpnfo_FindInDHCP = "The Find-InDHCP script will search for a computer name or dash separated MAC address in all the DHCP scopes in all the DHCP servers.  
It will return the IP address and IP lease (or IP reservation) information if found.  This can be used to troubleshoot systems that are online, but not pingable by name, or to find the 
last known location of a system that is offline if its IP lease has not expired yet.
"
$fpnfo_FixSccmClient = "Syntax:  'Fix-SCCMClient Computername'  This script will remotely set the MECM site code (NET, SET or GET), delete the MECM client certificate, 
restart the ccmexec.exe service causing the client to pull a fresh certificate from the MECM server, and then run all of the client actions that are available under the control panel configuration manager
actions tab.  This should cause the client to check in with the MECM server and start pulling the updates.
"

$fpnfo_GetTestGrpApps = " Get-TestGrpApps 'application name or partial name' will reach out to each of the computers in the MECM test group and pull the installed application 
data for the application name you specify.  Handy if you are testing an application deployment.

The 'DisplayVersion' attribute will tell you if the version that is being pushed by MECM is the version that is actually installed.

For example 'Get-TestGrpApps Firefox' will return the installation data for any application containing the word 'Firefox'.  

If you use the function to search for an application that has spaces in the name put quotes around the name (get-TestGrpApps 'Mozilla Firefox').

The test group is hard-coded in the script, so if we add or remove computers from the test group, we will need to modify the powershell script and run update-sysadmin. 
"

$fpnfo_Laplocker = "The command 'Lapslocker' will open a GUI where you can type a computername and look up either the LAPS password or the Bitlocker Recovery Key for the system. 

If you don't have access to the attribute where the data is stored you will get an error message. 

"
$fpnfo_MigratePrinters = " Migrates printers from one print server to another.  Prompts for 'From' and 'To'

"

$fpnfo_ReportError = "
    .Synopsis
  Report an error OR bug that is being caused by any script or module created by the SysAdmins Script Team. 

 .Description
  Allows the Administrator to create a report that describes an error or bug pertaining to a script and/or module created by the SysAdmins Script Team.

 .Example
    Report-Error
    # Runs the Script with the standard protocols.
"

$fpnfo_UpdateSysAdmin = "The Update-SysAdmin script is used solely for updating the powershell module that allows you to utilize all these wonderful commands that make your life easier. This command shouldn't need to be ran by
anyone other than the team members who work on the module and scripts. The command has been updated to automatically update all machines whenever a change is made to the module. 

NOTE : As of MAY 2023, the Test-ExchangeServerHealth script will NO LONGER run this command.

Additional Parameters;
-Single : Specify a single computer to update the module onto, and asks if you wish to add it to the automatic update list.
-Refresh : Runs powershell command to refresh the powershell session.
"

$fpnfo_GetUptimeReport = "Get-UptimeReport is designed to go through either DCs, Member Servers, or All servers and retrieve the 'uptime' on them. 
The uptime is shown as the last day and how many hours since the server has been restarted. If a server hasn't been rebooted for a long period of time, it could be possible that it is missing critical patches.

Additional Parameters;
-AllDCs : Goes through all the Domain Controllers
-MemberServers : Goes through all the member servers.
-AllServers : Goes through every known server.
"

$fpnfo_RunHashRefresh = "Run-HashRefresh refreshes the certificate hashes for users that use a PKI token to log in.  
It is run annually and will force any logged in users to log out and log in again
"

$fpnfo_GetCertRequest = "Get-CertRequest automates the process of generating a certificate request for a system.  Syntax is 'Get-CertRequest Systemname'.  
A GUI will appear that has all the information that must copied and pasted into the certificate request website.
"

$fpnfo_RunDataCall = "Run-DataCall
.Synopsis
  Automates any type of Datacall for users and computers 

 .Description
  Allows admins to quickly create datacalls for any user(s) and/or any machine(s)

 .Example
   Run-DataCall
   # Load the DataCall Script
"

$fpnfo_CheckSvcLastPassSet = "Check-SvcLastPassSet checks the domain *SVC* accounts and the top level domain admin account for password age compliance.  
Will flag any accounts that have a password more than 350 days old.  The list of non-compliant accounts should be sent to Cyber to coordinate the password
change with the account owner.
"

$fpnfo_InstallPowerCLI = "InstallPowerCLI installs VMware.PowerCLI powershell module.  
With this powershell module, Virtual machines and hosts can be accessed from the powershell console which is sometimes preferable to the web-based Vsphere interface.
"

$fpnfo_CheckSMTP = "CheckSMTP Requires VMware.PowerCLI.  It launches a Open-VMConsoleWindow to each SMTP server so that the admin can log in and check the queues.  
Should be run each shift. 
"




 

Function Add-Label{
param([int]$xsize,[int]$ysize,[int]$xpos,[int]$ypos,[string]$text,[switch]$fixed3d)

$lname = New-Object $Label
$lname.Size = New-Object $Size($xsize,$ysize)
$lname.Location = New-Object $Location($xpos,$ypos)
$lname.Text = $text
If($fixed3d){$lname.BorderStyle = 'Fixed3D'}

$lname
}

Function Add-Button{
param([int]$xsize,[int]$ysize,[int]$xpos,[int]$ypos,[string]$text)

$bname = New-Object $Button
$bname.Size = New-Object $Size($xsize,$ysize)
$bname.Location = New-Object $Location($xpos,$ypos)
$bname.Text = $text

$bname
}

Function Add-Background{
param([int]$xsize,[int]$ysize,[int]$xpos,[int]$ypos,[string]$color)

$bkgname = New-Object $PictureBox
$bkgname.Size = New-Object $Size($xsize,$ysize)
$bkgname.Location = New-Object $Location($xpos,$ypos)
$bkgname.BackColor = $color
$bkgname.BorderStyle = 'Fixed3D'

$bkgname
}

Function Add-TextBox{
param([int]$xsize,[int]$ysize,[int]$xpos,[int]$ypos)

$cbname = New-Object $RichTextBox
$cbname.Size = New-Object $Size($xsize,$ysize)
$cbname.Location = New-Object $Location($xpos,$ypos)
$cbname.Font = $SecondaryFontStyle

$cbname
}

Function Add-ListBox{
param([int]$xsize,[int]$ysize,[int]$xpos,[int]$ypos)

$cbname = New-Object $ListBox
$cbname.Size = New-Object $Size($xsize,$ysize)
$cbname.Location = New-Object $Location($xpos,$ypos)
$cbname.Font = $SecondaryFontStyle

$cbname
}

Function Script-Info {
$fpnfoBox.Clear()
If($ScriptList.SelectedIndex -eq 0){$fpnfoBox.Text = $fpnfo_GetDomainPrinters ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 1){$fpnfoBox.Text = $fpnfo_GetExpiredCerts ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 2){$fpnfoBox.Text = $fpnfo_GetStaleUsers ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 3){$fpnfoBox.Text = $fpnfo_GetStaleAdmins ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 4){$fpnfoBox.Text = $fpnfo_GetUptimeReport ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 5){$fpnfoBox.Text = $fpnfo_DisableStaleUsers ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 6){$fpnfoBox.Text = $fpnfo_DisableStaleAdmins ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 7){$fpnfoBox.Text = $fpnfo_ClonePermissions ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 8){$fpnfoBox.Text = $fpnfo_ResetBitLockerKey ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 9){$fpnfoBox.text = $fpnfo_checkdrivespace ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 10){$fpnfoBox.Text = $fpnfo_TestExchangeServerHealth ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 11){$fpnfoBox.Text = $fpnfo_UpdateSysAdmin ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 12){$fpnfoBox.Text = $fpnfo_FindInDHCP ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 13){$fpnfoBox.Text = $fpnfo_FixSccmClient ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 14){$fpnfoBox.Text = $fpnfo_GetTestGrpApps ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 15){$fpnfoBox.Text = $fpnfo_Laplocker ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 16){$fpnfoBox.Text = $fpnfo_MigratePrinters ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 17){$fpnfoBox.Text = $fpnfo_ReportError ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 18){$fpnfoBox.Text = $fpnfo_RunHashRefresh ; netcolor -Lowside -Snet}
If($ScriptList.SelectedIndex -eq 19){$fpnfoBox.Text = $fpnfo_GetCertRequest ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 20){$fpnfoBox.Text = $fpnfo_RunDataCall ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 21){$fpnfoBox.Text = $fpnfo_CheckSvcLastPassSet ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 22){$fpnfoBox.Text = $fpnfo_InstallPowerCLI ; netcolor -Lowside -Snet -TSnet}
If($ScriptList.SelectedIndex -eq 23){$fpnfoBox.Text = $fpnfo_CheckSMTP ; netcolor -Lowside -Snet -TSnet}

}

Function netcolor {
Param ([switch]$Lowside, [switch]$Snet, [switch]$TSnet)

If($Lowside){$lbl_Lowside.Visible = $true}Else{$lbl_Lowside.Visible = $false}
If($Snet){$lbl_Snet.Visible = $true}Else{$lbl_Snet.Visible = $false}
IF($TSnet){$lbl_TSnet.Visible = $true}Else{$lbl_TSnet.Visible = $false}

}


$HelpForm = New-Object $Form ; $HelpForm.Size = New-Object $Size(1000,600)
$ScriptList = Add-ListBox -xsize 250 -ysize 410 -xpos 20 -ypos 20 ; $ScriptList.DataSource = $ActiveScripts ; $ScriptList.Add_SelectedIndexChanged{Script-Info}
$fpnfoBox = Add-TextBox -xsize 650 -ysize 500 -xpos 300 -ypos 20 ; $fpnfoBox.ReadOnly = $true ; $fpnfoBox.Multiline = $true

$lbl_Network = Add-Label -xsize 195 -ysize 25 -xpos 20 -ypos 440 -text "Available on :" -fixed3d ; $lbl_Network.Font = $PrimaryFontStyle ; $lbl_Network.BackColor = 'LightGray'
$lbl_Lowside = Add-Label -xsize 20 -ysize 22 -xpos 140 -ypos 441 -text "N"  ; $lbl_Lowside.Font = $networkfont ; $lbl_Lowside.ForeColor = 'Green' ; $lbl_Lowside.BackColor = 'Green'
$lbl_Snet = Add-Label -xsize 20 -ysize 22 -xpos 165 -ypos 441 -text "S"  ; $lbl_Snet.Font = $networkfont ; $lbl_Snet.ForeColor = 'Red' ; $lbl_Snet.BackColor = 'Red'
$lbl_TSnet = Add-Label -xsize 20 -ysize 22 -xpos 190 -ypos 441 -text "TS"  ; $lbl_TSnet.Font = $networkfont ; $lbl_TSnet.ForeColor = 'Orange' ; $lbl_TSnet.BackColor = 'Orange'




$Objects = @($ScriptList,$fpnfoBox,$lbl_Network,$lbl_Lowside,$lbl_Snet,$lbl_TSnet)
$HelpForm.Controls.AddRange($Objects)
 $lbl_Network.SendToBack()
[System.Windows.Forms.Application]::Run($HelpForm)
}

#########################################################
