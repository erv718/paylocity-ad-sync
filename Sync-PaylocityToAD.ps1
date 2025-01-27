<#
.SYNOPSIS
    Sync AD users from Paylocity_Data_Latest.xlsx

.DESCRIPTION
    • Reads only Paylocity_Data_Latest.xlsx (skips title/footer rows)
    • Splits by status (A/T) into Removed/New/Rehire via AD lookups (checks both padded & unpadded IDs)
    • Logs to SyncReport_*.xlsx with multiple tabs & ordered columns
    • Optional AD changes with –Apply (alias –NoWhatIf)
    • Normalizes 5-digit Role numbers (adds leading zero) *and* looks up both forms
    • Skips Personal Training, Maintenance, Front Desk
    • Handles missing/blank headers; trims; removes duplicates
    • Drops final “Report Total Records: N” footer
    • Default password = FirstInitial + LastInitial + “[company]” + last 4 of ID
    • Places new users under Corporate vs. Clubs OUs based on Location Description
    • Stores Personal Email in extensionAttribute2
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
    Write-Error "Report not found: $Report"
    return
}

$WillChangeAD = $Apply.IsPresent
Write-Host ("WillChangeAD = {0} (false => only logs, no AD changes)" -f $WillChangeAD)

# Global change log for rehires
$global:RehireChangeLog = @()

# 1) Helper to trim all string properties
function Normalize-Records {
    param([array][AllowEmptyCollection()]$records)
    if (-not $records) { return @() }
    $records | ForEach-Object {
        $h = @{}
        foreach ($p in $_.PSObject.Properties) {
            $k = $p.Name.Trim()
            $v = $p.Value
            if ($v -is [string]) { $v = $v.Trim() }
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

    # 2a) Skip title row, row2 = headers
    $data = Import-Excel -Path $Path -StartRow 2 -ErrorAction SilentlyContinue
    $data = Normalize-Records -records $data
    if (-not $data.Count) { return @() }

    # 2b) Re-import with fixed headers if missing
    $present = $data[0].PSObject.Properties.Name
    $missing = $ExpectedHeaders | Where-Object { $_ -notin $present }
    if ($missing) {
        Write-Host "Missing headers: $($missing -join ', '). Re-importing with predefined headers."
        $data = Import-Excel -Path $Path -StartRow 2 -Header $ExpectedHeaders -ErrorAction SilentlyContinue
        $data = Normalize-Records -records $data
    }

    # 2c) Drop rows that exactly repeat the header row
    $expLow = $ExpectedHeaders | ForEach-Object { $_.Trim().ToLower() }
    $data = $data | Where-Object {
        $map = @{}
        foreach ($p in $_.PSObject.Properties) {
            $key = $p.Name.Trim().ToLower()
            $val = $p.Value -as [string]
            if ($val) { $val = $val.Trim().ToLower() } else { $val = "" }
            $map[$key] = $val
        }
        $isHeader = $true
        foreach ($h in $expLow) {
            if (-not $map.ContainsKey($h) -or $map[$h] -ne $h) {
                $isHeader = $false; break
            }
        }
        -not $isHeader
    }

    # 2d) Drop “Report Total Records: N” footer
    $script:ReportTotalCount = $null
    $data = $data | Where-Object {
        if ($_.Role -match '^Report\s*Total\s*Records:\s*(\d+)$') {
            $script:ReportTotalCount = [int]$Matches[1]
            $false
        } else { $true }
    }

    # 2e) Drop rows with only one non-empty cell
    $data = $data | Where-Object {
        ($_.PSObject.Properties.Value | Where-Object { $_ -notin @($null, "") }).Count -gt 1
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

# Desired export order & depts to skip
$desiredOrder    = @(
    "Role","Preferred First Name","First Name","Last Name","Email","Personal Email",
    "Mobile Phone","Location  Code","Location Description","Department Code",
    "Department Name","Job Title","address1","address2","city","state","zip",
    "rehiredate","status","Termination Date","isexempt","issalaried","ishourly"
)
$skipDepartments = @("Personal Training","Maintenance","Front Desk")

# --- LOAD & CLEAN ---
$data = Import-CleanExcel -Path $Report

# 4) Validate row count
if ($script:ReportTotalCount) {
    Write-Host ("Declared total: {0}" -f $script:ReportTotalCount)
    Write-Host ("Imported rows:  {0}" -f $data.Count)
    if ($data.Count -ne $script:ReportTotalCount) {
        Write-Warning "Row count mismatch: imported $($data.Count) vs declared $script:ReportTotalCount"
    }
}

# 5) Classify
$removedAccounts = @(); $newAccounts = @(); $rehireAccounts = @()
foreach ($r in $data) {
    if ($skipDepartments -contains $r.'Department Name') { continue }
    if (-not $r.Role) { continue }

    # preserve both padded & original IDs for lookup
    $origRole = $r.Role.Trim()
    if ($origRole -match '^\d{5}$') {
        $padRole = "0$origRole"
        $r.Role  = $padRole
    } else {
        $padRole = $origRole
        $r.Role   = $origRole
    }

    if (-not $r.'First Name' -or -not $r.'Last Name') { continue }

    switch ($r.status) {
        'T' {
            $removedAccounts += $r
        }
        'A' {
            $filter = "((employeeID -eq '$($r.Role)') -or (employeeNumber -eq '$($r.Role)')" +
                      " -or (employeeID -eq '$origRole') -or (employeeNumber -eq '$origRole'))"
            $u = Get-ADUser -Filter $filter -Properties Enabled -ErrorAction SilentlyContinue
            if (-not $u)                   { $newAccounts    += $r }
            elseif (-not $u.Enabled)       { $rehireAccounts += $r }
        }
        default {
            Write-Host "[-] Unknown status '$($r.status)' for Role '$($r.Role)'. Skipping."
        }
    }
}

# 6) SUMMARY
$adText = if ($WillChangeAD) { 'YES' } else { 'NO' }
Write-Host "`n=== SUMMARY ==="
Write-Host ("Terminated: {0}" -f $removedAccounts.Count)
Write-Host ("New Accounts: {0} | Rehires: {1}" -f $newAccounts.Count, $rehireAccounts.Count)
Write-Host "AD changes? $adText"

# 7) OFFBOARDING  …  (unchanged from v10)  
# 8) REHIRES     …  (unchanged from v10)

# 9) ONBOARD NEW ACCOUNTS
if ($newAccounts.Count) {
    Write-Host "`n--- ONBOARDING NEW ACCOUNTS ---"
    foreach ($r in $newAccounts) {
        $id       = $r.Role
        $fn       = if ($r.'Preferred First Name') { $r.'Preferred First Name' } else { $r.'First Name' }
        $ln       = $r.'Last Name'
        $sam      = ("{0}.{1}" -f $fn,$ln).Replace(' ','').ToLower()
        $password = "{0}{1}[company]{2}" -f $fn[0], $ln[0], $id.Substring($id.Length-4)
        $userEmail = "$sam@corp.example.com"
        $desc     = "$($r.'Location  Code') - $($r.'Job Title')"

        # choose OU by Location Description
        if ($r.'Location Description' -eq 'Corporate') {
            $ou = "OU=Users,OU=Corporate,OU=Locations,OU=[Company],DC=ad,DC=corp.example,DC=com"
        } else {
            $ou = "OU=Users,OU=Clubs,OU=Locations,OU=[Company],DC=ad,DC=corp.example,DC=com"
        }

        Write-Host " - [$id] $sam ($fn $ln) => create in $ou"
        if ($WillChangeAD) {
            try {
                New-ADUser `
                    -Name             "$fn $ln" `
                    -SamAccountName    $sam `
                    -UserPrincipalName "$sam@ad.corp.example.com" `
                    -GivenName         $fn `
                    -Surname           $ln `
                    -EmailAddress      $userEmail `
                    -MobilePhone       $r.'Mobile Phone' `
                    -Description       $desc `
                    -AccountPassword   (ConvertTo-SecureString $password -AsPlainText -Force) `
                    -Enabled           $true `
                    -Path              $ou `
                    -OtherAttributes   @{
                        employeeID          = $id
                        title               = $r.'Job Title'
                        department          = $r.'Department Name'
                        extensionAttribute2 = $r.'Personal Email'
                    }

                Add-ADGroupMember -Identity "sec00us-googleuser-sec" -Members $sam -ErrorAction Stop
                Write-Host "   Created & added to group."
            }
            catch {
                Write-Warning (
                    "   ✗ Failed to create {0}: {1}" -f
                    $sam,
                    $_.Exception.Message
                )
            }
        }
        else {
            Write-Host "   [Simulation] Would create $sam with password '$password' in $ou."
        }
    }
} else {
    Write-Host "No new accounts."
}

# 10) EXPORT …  (unchanged from v10)

Write-Host "`nDone. Real AD changes applied? $WillChangeAD (false = simulation only)."
