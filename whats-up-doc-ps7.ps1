Import-module Active-Directory

function Whats-Up-Doc-ps7a {
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

$serverlist = $serverlist | sort -Unique
$offline = $serverlist | ForEach-Object -Parallel {
    $server = $_ 

    if (test-connection -ComputerName $server -count 4 -Quiet) {
        Write-Host "$server is up"
    }
    else { 
        Write-Host "$server is down" -f Red
        $server
    }
} -ThrottleLimit 100


Write-host ""
Write-host "The Following Systems are down as of $(get-date) $([system.timezoneinfo]::Local.id):" -ForegroundColor Red
$body = "`n Technician: $(whoami)"
$body += "`n Refer to the High Priority Sites Memo on TS Sharepoint Portal for reporting responsibilites."
$body += "`n "
$body += "`n The following systems are down as of $(get-date) $([system.timezoneinfo]::Local.id):"

$offlineInfo = $offline | ForEach-Object -Parallel {
    $off = $_

    try {
        [string]$canname = (Get-ADComputer -Identity $off -Properties CanonicalName -ErrorAction Si | select -ExpandProperty CanonicalName)

        if ($canname) {
            $canname = $canname -replace "newton.pentagon.mil/Sites/",""
            $canname = $canname -replace [regex]::Escape($off), ""

            [pscustomobject] @{
                Name          = $off
                CanonicalName = $canname
            }
        }
        else {
            [pscustomobject] @{
                Name          = $off 
                CanonicalName = "No Canonical Name in AD"
            }
        }
    }
    catch {
        [pscustomobject] @{
            Name          = $off 
            CanonicalName = "No Canonical Name in AD"
        }
    }

    $offlineInfo = $offlineInfo | sort Name

    foreach ($item in $offlineInfo) {
        Write-Host "$($item.Name) - $($item.CanonicalName)" -f Red 
        $body += "`n$($item.Name) - $($item.CanonicalName)"
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

Whats-Up-Doc-ps7a @PSBoundParameters