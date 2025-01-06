<#
.SYNOPSIS
    Sync AD users from Paylocity_Data_Latest.xlsx

.DESCRIPTION
    • Reads only Paylocity_Data_Latest.xlsx (skips title/footer rows)
    • Splits by status (A/T) into Removed/New/Rehire via AD lookups
    • Logs to SyncReport_*.xlsx with multiple tabs & ordered columns
    • Optional AD changes with –Apply (alias –NoWhatIf)
    • Normalizes 5-digit Role numbers (adds leading zero)
    • Skips Personal Training, Maintenance, Front Desk
    • Handles missing/blank headers; trims; removes duplicates
    • Drops final “Report Total Records: N” footer
    • Default password = FirstInitial + LastInitial + “[company]” + last 4 of ID
#>

[CmdletBinding()]
param(
    [string]$Report = "C:\Reports\Paylocity_Data_Latest.xlsx",
    [Alias("NoWhatIf")][switch]$Apply
)

# 0) Modules & basic checks
Import-Module ActiveDirectory -ErrorAction Stop
Import-Module ImportExcel     -ErrorAction Stop

Write-Host "Report => $Report"
if (-not (Test-Path $Report)) {
    Write-Error "Report not found: $Report"; return
}

$WillChangeAD = $Apply.IsPresent
Write-Host ("WillChangeAD = {0} (false => only logs, no AD changes)" -f $WillChangeAD)

# Global rehire-change log
$global:RehireChangeLog = @()

# 1) Normalize trim helper
function Normalize-Records {
    param([array][AllowEmptyCollection()]$records)
    if (-not $records) { return @() }
    $records | ForEach-Object {
        $h = @{}
        foreach ($p in $_.PSObject.Properties) {
            $k = $p.Name.Trim()
            $v = $p.Value
            if ($v -and ($v -is [string])) { $v = $v.Trim() }
            $h[$k] = $v
        }
        [PSCustomObject]$h
    }
}

# 2) Import & clean Excel
function Import-CleanExcel {
    param(
        [string]$Path,
        [string[]]$ExpectedHeaders = @(
            "Role","Email","Personal Email","First Name","Preferred First Name","Last Name",
            "Mobile Phone","Location  Code","Location Description","Department Code",
            "Department Name","Job Title","Salary or Hourly","Original Hire Date","DOB",
            "Type","Training Phase","address1","address2","city","state","zip",
            "rehiredate","status","Termination Date","isexempt","issalaried","ishourly"
        )
    )

    # 2a) Skip title row 1, treat row2 as headers
    $data = Import-Excel -Path $Path -StartRow 2 -ErrorAction SilentlyContinue
    $data = Normalize-Records -records $data
    if (-not $data.Count) { return @() }

    # 2b) If headers missing, re-import with our list
    $present = $data[0].PSObject.Properties.Name
    $missing = $ExpectedHeaders | Where-Object { $_ -notin $present }
    if ($missing) {
        Write-Host "Missing headers: $($missing -join ', '). Re-importing with predefined headers."
        $data = Import-Excel -Path $Path -StartRow 2 -Header $ExpectedHeaders -ErrorAction SilentlyContinue
        $data = Normalize-Records -records $data
    }

    # 2c) Drop duplicate-header rows
    $expLow = $ExpectedHeaders | % { $_.Trim().ToLower() }
    $data = $data | Where-Object {
        $map = @{}
        foreach ($p in $_.PSObject.Properties) {
            $k = $p.Name.Trim().ToLower()
            $v = ($p.Value -as [string])
            $v = if ($v) { $v.Trim().ToLower() } else { "" }
            $map[$k] = $v
        }
        $isHdr = $true
        foreach ($h in $expLow) {
            if (-not $map.ContainsKey($h) -or $map[$h] -ne $h) {
                $isHdr = $false; break
            }
        }
        -not $isHdr
    }

    # 2d) Drop “Report Total Records: N” footer
    $script:ReportTotalCount = $null
    $data = $data | Where-Object {
        if ($_.Role -and ($_.Role -match '^Report\s*Total\s*Records:\s*(\d+)$')) {
            $script:ReportTotalCount = [int]$Matches[1]
            $false
        } else { $true }
    }

    # 2e) Drop stray single-cell rows
    $data = $data | Where-Object {
        ($_.PSObject.Properties.Value | Where-Object { $_ -ne $null -and $_ -ne "" }).Count -gt 1
    }

    return $data
}

# 3) Re-order for Excel export
function Order-Columns {
    param(
        [array][AllowEmptyCollection()]$data,
        [string[]]$Order
    )
    if (-not $data.Count) { return @() }
    $have = $data[0].PSObject.Properties.Name
    $cols = $Order | Where-Object { $have -contains $_ }
    return $data | Select-Object -Property $cols
}

# Desired column order
$desiredOrder = @(
    "Role","Preferred First Name","First Name","Last Name",
    "Email","Personal Email","Mobile Phone",
    "Location  Code","Location Description","Department Code",
    "Department Name","Job Title","address1","address2",
    "city","state","zip","rehiredate","status","Termination Date",
    "isexempt","issalaried","ishourly"
)

# Skip these depts
$skipDepartments = @("Personal Training","Maintenance","Front Desk")

# --- LOAD & CLEAN ---
$data = Import-CleanExcel -Path $Report

# 4) Validate row count
if ($script:ReportTotalCount) {
    Write-Host ("Declared total: {0}" -f $script:ReportTotalCount)
    Write-Host ("Imported rows:  {0}" -f $data.Count)
    if ($data.Count -ne $script:ReportTotalCount) {
        Write-Warning -Message "Row count mismatch: imported $($data.Count) vs declared $script:ReportTotalCount"
    }
}

# 5) Classify
$removedAccounts = @(); $newAccounts = @(); $rehireAccounts = @()
foreach ($r in $data) {
    if ($skipDepartments -contains $r.'Department Name') { continue }
    if (-not $r.Role) { continue }
    if ($r.Role.Length -eq 5) { $r.Role = "0$r.Role" }
    if ([string]::IsNullOrWhiteSpace($r.'First Name') -or [string]::IsNullOrWhiteSpace($r.'Last Name')) { continue }

    switch ($r.status) {
        'T' { $removedAccounts   += $r }
        'A' {
            $f  = "(employeeID -eq '$($r.Role)') -or (employeeNumber -eq '$($r.Role)')"
            $u  = Get-ADUser -Filter $f -Properties Enabled -ErrorAction SilentlyContinue
            if (-not $u)             { $newAccounts    += $r }
            elseif (-not $u.Enabled) { $rehireAccounts += $r }
        }
        default {
            Write-Host "[-] Unknown status '$($r.status)' for Role $($r.Role). Skipping."
        }
    }
}

# 6) SUMMARY
Write-Host "`n=== SUMMARY ==="
Write-Host ("Terminated: {0}" -f $removedAccounts.Count)
Write-Host ("New Accounts: {0} | Rehires: {1}" -f $newAccounts.Count, $rehireAccounts.Count)
$adText = if ($WillChangeAD) { "YES" } else { "NO" }
Write-Host ("AD changes? {0}" -f $adText)

# 7) OFFBOARDING
if ($removedAccounts.Count) {
    Write-Host "`n--- OFFBOARDING ---"
    foreach ($r in $removedAccounts) {
        $id = $r.Role; $fn = $r.'First Name'; $ln = $r.'Last Name'
        Write-Host " - [$id] $fn $ln => disable/move in AD"

        $all = Get-ADUser -Filter "(employeeID -eq '$id') -or (employeeNumber -eq '$id')" `
              -Properties SamAccountName,DistinguishedName,Enabled -ErrorAction SilentlyContinue
        $act = $all | Where-Object { $_.Enabled -eq $true }

        switch ($act.Count) {
            1 {
                $u = $act[0]
                if ($WillChangeAD) {
                    Disable-ADAccount -Identity $u.SamAccountName
                    Move-ADObject     -Identity $u.DistinguishedName `
                                      -TargetPath "OU=Users,OU=Disabled,OU=[Company],DC=ad,DC=corp.example,DC=com"
                    $groups = Get-ADUser -Identity $u.SamAccountName -Properties MemberOf |
                              Select-Object -ExpandProperty MemberOf
                    foreach ($dn in $groups) {
                        if ($dn -notlike "*Domain Users*") {
                            Remove-ADGroupMember -Identity $dn -Members $u.SamAccountName -Confirm:$false
                        }
                    }
                    Write-Host "   Disabled & moved."
                } else {
                    Write-Host "   [Simulation] Would disable & move $($u.SamAccountName)."
                }
            }
            0 {
                Write-Warning -Message "No active AD user found for Role $id - skipping off-boarding"
            }
            default {
                Write-Warning -Message "Multiple active matches ($($act.Count)) for Role $id - skipping off-boarding"
            }
        }
    }
} else {
    Write-Host "No terminated users."
}

# 8) REHIRES
if ($rehireAccounts.Count) {
    Write-Host "`n--- REHIRES ---"
    foreach ($r in $rehireAccounts) {
        $id = $r.Role; $fn = $r.'First Name'; $ln = $r.'Last Name'
        Write-Host " - [$id] $fn $ln => re-enable & update"

        $f   = "(employeeID -eq '$id') -or (employeeNumber -eq '$id')"
        $usr = Get-ADUser -Filter $f -Properties Enabled,givenName,sn,title,department,mail `
               -ErrorAction SilentlyContinue
        if (-not $usr) {
            Write-Warning -Message "AD user not found for rehire."
            continue
        }

        if ($WillChangeAD) {
            $before = Get-ADUser -Identity $usr.SamAccountName -Properties givenName,sn,title,department,mail
            $newH   = @{
                givenName  = $fn
                sn         = $ln
                title      = $r.'Job Title'
                department = $r.'Department Name'
                mail       = $r.Email
            }
            foreach ($k in $newH.Keys) {
                if ($before.$k -ne $newH[$k]) {
                    $global:RehireChangeLog += [PSCustomObject]@{
                        SamAccountName = $usr.SamAccountName
                        Attribute      = $k
                        Before         = $before.$k
                        After          = $newH[$k]
                    }
                }
            }
            Enable-ADAccount -Identity $usr.SamAccountName
            Set-ADUser       -Identity $usr.SamAccountName -Replace $newH
            Write-Host "   Re-enabled & updated."
        } else {
            Write-Host "   [Simulation] Would re-enable & update."
        }
    }
} else {
    Write-Host "No rehires."
}

# 9) ONBOARD NEW ACCOUNTS
if ($newAccounts.Count) {
    Write-Host "`n--- ONBOARDING NEW ACCOUNTS ---"
    foreach ($r in $newAccounts) {
        $id    = $r.Role
        $fn    = if ($r.'Preferred First Name') { $r.'Preferred First Name' } else { $r.'First Name' }
        $ln    = $r.'Last Name'
        $sam   = ("{0}.{1}" -f $fn,$ln).Replace(' ','').ToLower()
        $email = "$sam@corp.example.com"
        $desc  = "$($r.'Location  Code') - $($r.'Job Title')"
        $pw    = "{0}{1}[company]{2}" -f $fn[0], $ln[0], $id.Substring($id.Length-4)

        Write-Host " - [$id] $sam ($fn $ln) => create"
        if ($WillChangeAD) {
            New-ADUser -Name             "$fn $ln" `
                       -SamAccountName    $sam `
                       -UserPrincipalName "$sam@ad.corp.example.com" `
                       -GivenName         $fn `
                       -Surname           $ln `
                       -EmailAddress      $email `
                       -MobilePhone       $r.'Mobile Phone' `
                       -Description       $desc `
                       -AccountPassword   (ConvertTo-SecureString $pw -AsPlainText -Force) `
                       -Enabled           $true `
                       -Path              "OU=Users,DC=ad,DC=corp.example,DC=com" `
                       -OtherAttributes   @{
                           employeeID          = $id
                           title               = $r.'Job Title'
                           department          = $r.'Department Name'
                           extensionAttribute2 = $r.'Personal Email'
                       }
            Add-ADGroupMember -Identity "sec00us-googleuser-sec" -Members $sam -ErrorAction SilentlyContinue
            Write-Host "   Created & added to group."
        } else {
            Write-Host "   [Simulation] Would create with password '$pw'."
        }
    }
} else {
    Write-Host "No new accounts."
}

# 10) EXPORT
$ts = Get-Date -Format "yyyy-MM-dd_HH-mm"
$out = "C:\Reports\SyncReport_$ts.xlsx"
Write-Host "`n--- Exporting to $out ---"

Order-Columns -data $removedAccounts -Order $desiredOrder |
    Export-Excel -Path $out -WorksheetName "Removed" -AutoSize -Title "Removed Users"

Order-Columns -data $newAccounts -Order $desiredOrder |
    Export-Excel -Path $out -WorksheetName "Added"   -AutoSize -Title "New Accounts" -Append

Order-Columns -data $rehireAccounts -Order $desiredOrder |
    Export-Excel -Path $out -WorksheetName "Rehires" -AutoSize -Title "Rehired/Existing" -Append

if ($global:RehireChangeLog.Count) {
    $global:RehireChangeLog |
        Export-Excel -Path $out -WorksheetName "RehireChanges" -AutoSize -Title "Before vs After" -Append
} else {
    @() | Export-Excel -Path $out -WorksheetName "RehireChanges" -Title "No Rehire Changes" -Append
}

Write-Host "`nDone. Real AD changes applied? $WillChangeAD (false = simulation only)."  
