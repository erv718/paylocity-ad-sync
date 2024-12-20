<#
    Single-file Paylocity sync script:
      • Reads only Paylocity_Data_Latest.xlsx (skipping the blank first row)
      • Uses status “A” vs “T” to split into Removed/New/Rehire via AD lookups
      • Logs to SyncReport_*.xlsx with multiple tabs & ordered columns
      • Optional AD changes with –Apply (alias –NoWhatIf)
      • Normalizes 5-digit Role numbers (adds leading zero)
      • Skips departments: Personal Training, Maintenance, Front Desk
      • Handles missing/blank headers with a predefined list
      • Trims headers & values; removes actual duplicate header rows
      • Captures & drops the final “Report Total Records: N” footer
      • Drops any row with only one non-empty cell (the stray “1”/“2”)
      • Generates default password = FirstInitial + LastInitial + "[company]" + last 4 of ID
#>

[CmdletBinding()]
param(
    [string]$Report = "C:\Reports\Paylocity_Data_Latest.xlsx",
    [Alias("NoWhatIf")][switch]$Apply
)

# Import modules
Import-Module ActiveDirectory -ErrorAction Stop
Import-Module ImportExcel     -ErrorAction Stop

Write-Host "Report => $Report"
if (-not (Test-Path $Report)) {
    Write-Error "Report not found: $Report"
    return
}

$WillChangeAD = $Apply.IsPresent
Write-Host "WillChangeAD = $WillChangeAD (false => only logs, no AD changes)"

# Global rehire-change log
$global:RehireChangeLog = @()

# -- Normalize any array of PSCustomObjects, trimming keys & string values
function Normalize-Records {
    param(
        [Parameter(Mandatory=$true)]
        [array][AllowEmptyCollection()]$records
    )
    if (-not $records) { return @() }
    return $records | ForEach-Object {
        $o = @{}
        foreach ($p in $_.PSObject.Properties) {
            $k = $p.Name.Trim()
            $v = $p.Value
            if ($v -and ($v -is [string])) { $v = $v.Trim() }
            $o[$k] = $v
        }
        [PSCustomObject]$o
    }
}

# -- Load & clean the Excel sheet
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

    # 1) Skip the blank first row; treat row 2 as header
    $data = Import-Excel -Path $Path -StartRow 2 -ErrorAction SilentlyContinue
    if (-not $data) { return @() }
    $data = Normalize-Records -records $data

    # 2) If any expected header is missing, force a re-import with our list
    $present = $data[0].PSObject.Properties.Name
    $missing = $ExpectedHeaders | Where-Object { $_ -notin $present }
    if ($missing) {
        Write-Host "Missing headers: $($missing -join ', '). Re-importing with predefined headers."
        $data = Import-Excel -Path $Path -StartRow 2 -Header $ExpectedHeaders -ErrorAction SilentlyContinue
        $data = Normalize-Records -records $data
    }

    # 3) Remove *actual* duplicate-header rows (where every cell equals its own header)
    $expectedLower = $ExpectedHeaders | ForEach-Object { $_.Trim().ToLower() }
    $data = $data | Where-Object {
        $rowMap = @{}
        foreach ($p in $_.PSObject.Properties) {
            $key   = $p.Name.Trim().ToLower()
            $value = ($p.Value -as [string])
            $value = if ($value) { $value.Trim().ToLower() } else { "" }
            $rowMap[$key] = $value
        }
        # Detect header-row if for *all* expectedLower: rowMap[h]==h
        $isHeaderRow = $true
        foreach ($h in $expectedLower) {
            if (-not $rowMap.ContainsKey($h) -or $rowMap[$h] -ne $h) {
                $isHeaderRow = $false
                break
            }
        }
        -not $isHeaderRow
    }

    # 4) Capture & drop final “Report Total Records: N” footer
    $script:ReportTotalCount = $null
    $data = $data | Where-Object {
        if ($_.Role -and ($_.Role -match '^Report\s*Total\s*Records:\s*(\d+)$')) {
            $script:ReportTotalCount = [int]$Matches[1]
            $false
        } else { $true }
    }

    # 5) Drop any stray row that has only one non-empty cell (the “1” or “2”)
    $data = $data | Where-Object {
        ($_.PSObject.Properties.Value |
         Where-Object { $_ -ne $null -and $_ -ne "" }).Count -gt 1
    }

    return $data
}

# -- Reorder columns for export, accepts empty collections
function Order-Columns {
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyCollection()]
        [array]$data,
        [Parameter(Mandatory=$true)]
        [string[]]$Order
    )
    if (-not $data.Count) { return @() }
    $existing = $data[0].PSObject.Properties.Name
    $cols     = $Order | Where-Object { $existing -contains $_ }
    return $data | Select-Object -Property $cols
}

# Desired export column order
$desiredOrder = @(
    "Role","Preferred First Name","First Name","Last Name",
    "Email","Personal Email","Mobile Phone",
    "Location  Code","Location Description","Department Code",
    "Department Name","Job Title","address1","address2",
    "city","state","zip","rehiredate","status","Termination Date",
    "isexempt","issalaried","ishourly"
)

# Departments to skip
$skipDepartments = @("Personal Training","Maintenance","Front Desk")

# --- 1) Load & clean ---
$data = Import-CleanExcel -Path $Report

# --- 2) (Optional) Verify row count matches footer ---
if ($script:ReportTotalCount) {
    Write-Host ("Declared total: {0}" -f $script:ReportTotalCount)
    Write-Host ("Imported rows:  {0}" -f $data.Count)
    if ($data.Count -ne $script:ReportTotalCount) {
        Write-Warning "Row count mismatch: imported $($data.Count) vs total $script:ReportTotalCount"
    }
}

# --- 3) Classify into Removed / New / Rehire ---
$removedAccounts = @()
$newAccounts     = @()
$rehireAccounts  = @()

foreach ($rec in $data) {
    if ($skipDepartments -contains $rec.'Department Name') { continue }
    if (-not $rec.Role)                         { continue }
    if ($rec.Role.Length -eq 5) { $rec.Role = "0" + $rec.Role }
    if ([string]::IsNullOrWhiteSpace($rec.'First Name') -or
        [string]::IsNullOrWhiteSpace($rec.'Last Name')) { continue }

    switch ($rec.status) {
        'T' { $removedAccounts += $rec }
        'A' {
            $id  = $rec.Role
            $f   = "(employeeID -eq '$id') -or (employeeNumber -eq '$id')"
            $u   = Get-ADUser -Filter $f -Properties Enabled -ErrorAction SilentlyContinue
            if (-not $u)             { $newAccounts   += $rec }
            elseif (-not $u.Enabled) { $rehireAccounts += $rec }
        }
        default {
            Write-Host "[-] Unknown status '$($rec.status)' for Role $($rec.Role). Skipping."
        }
    }
}

# --- 4) Summary ---
Write-Host "`n=== SUMMARY ==="
Write-Host ("Terminated: {0}" -f $removedAccounts.Count)
Write-Host ("New Accounts: {0} | Rehires: {1}" -f $newAccounts.Count, $rehireAccounts.Count)
$adText = if ($WillChangeAD) { "YES" } else { "NO" }
Write-Host ("AD changes? {0}" -f $adText)

# --- 5) OFFBOARDING (Terminated) ---
if ($removedAccounts.Count) {
    Write-Host "`n--- OFFBOARDING ---"
    foreach ($r in $removedAccounts) {
        $id = $r.Role; $fn = $r.'First Name'; $ln = $r.'Last Name'
        Write-Host " - [$id] $fn $ln => disable/move in AD"
        if ($WillChangeAD) {
            $f   = "(employeeID -eq '$id') -or (employeeNumber -eq '$id')"
            $usr = Get-ADUser -Filter $f -Properties SamAccountName,DistinguishedName -ErrorAction SilentlyContinue
            if ($usr) {
                Disable-ADAccount -Identity $usr.SamAccountName
                Move-ADObject     -Identity $usr.DistinguishedName `
                    -TargetPath "OU=Users,OU=Disabled,OU=[Company],DC=ad,DC=corp.example,DC=com"
                $grps = Get-ADUser -Identity $usr.SamAccountName -Properties MemberOf |
                        Select-Object -ExpandProperty MemberOf
                foreach ($dn in $grps) {
                    if ($dn -notlike "*Domain Users*") {
                        Remove-ADGroupMember -Identity $dn -Members $usr.SamAccountName -Confirm:$false
                    }
                }
                Write-Host "   Disabled & moved."
            } else {
                Write-Host "   [Warning] AD user not found."
            }
        }
    }
} else {
    Write-Host "No terminated users."
}

# --- 6) REHIRES (Re-enable & update) ---
if ($rehireAccounts.Count) {
    Write-Host "`n--- REHIRES ---"
    foreach ($a in $rehireAccounts) {
        $id = $a.Role; $fn = $a.'First Name'; $ln = $a.'Last Name'
        Write-Host " - [$id] $fn $ln => re-enable & update"
        $f   = "(employeeID -eq '$id') -or (employeeNumber -eq '$id')"
        $usr = Get-ADUser -Filter $f -Properties Enabled,mail,title,department,givenName,surname `
                 -ErrorAction SilentlyContinue
        if ($usr) {
            if ($WillChangeAD) {
                $attrs  = 'givenName','surname','title','department','mail'
                $before = Get-ADUser -Identity $usr.SamAccountName -Properties $attrs
                $newH   = @{
                    givenName  = $fn
                    surname    = $ln
                    title      = $a.'Job Title'
                    department = $a.'Department Name'
                    mail       = $a.Email
                }
                foreach ($at in $attrs) {
                    if ($before.$at -ne $newH[$at]) {
                        $global:RehireChangeLog += [PSCustomObject]@{
                            SamAccountName = $usr.SamAccountName
                            Attribute      = $at
                            Before         = $before.$at
                            After          = $newH[$at]
                        }
                    }
                }
                Enable-ADAccount -Identity $usr.SamAccountName
                Set-ADUser       -Identity $usr.SamAccountName -Replace $newH
                Write-Host "   Re-enabled & updated."
            } else {
                Write-Host "   [Simulation] Would re-enable & update."
            }
        } else {
            Write-Host "   [Warning] AD user not found."
        }
    }
} else {
    Write-Host "No rehires."
}

# --- 7) ONBOARDING NEW ACCOUNTS ---
if ($newAccounts.Count) {
    Write-Host "`n--- ONBOARDING NEW ACCOUNTS ---"
    foreach ($a in $newAccounts) {
        $id    = $a.Role
        $first = $a.'Preferred First Name'; if (-not $first) { $first = $a.'First Name' }
        $last  = $a.'Last Name'
        $local = ($first + "." + $last) -replace "\s",""
        $sam   = $local.ToLower()
        $email = "$sam@corp.example.com"
        $desc  = "$($a.'Location  Code') - $($a.'Job Title')"
        $pw    = "{0}{1}[company]{2}" -f $first.Substring(0,1), $last.Substring(0,1), $id.Substring($id.Length-4)

        Write-Host " - [$id] $sam ($first $last) => create"
        if ($WillChangeAD) {
            New-ADUser -Name             "$first $last" `
                       -SamAccountName    $sam `
                       -UserPrincipalName "$sam@ad.corp.example.com" `
                       -GivenName         $first `
                       -Surname           $last `
                       -EmailAddress      $email `
                       -MobilePhone       $a.'Mobile Phone' `
                       -Description       $desc `
                       -AccountPassword   (ConvertTo-SecureString $pw -AsPlainText -Force) `
                       -Enabled           $true `
                       -Path              "OU=Users,DC=ad,DC=corp.example,DC=com" `
                       -OtherAttributes   @{
                           employeeID          = $id
                           mail                = $email
                           title               = $a.'Job Title'
                           department          = $a.'Department Name'
                           mobile              = $a.'Mobile Phone'
                           extensionAttribute2 = $a.'Personal Email'
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

# --- 8) Export to Excel ---
$timeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
$excelFile = "C:\Reports\SyncReport_$timeStamp.xlsx"
Write-Host "`n--- Exporting to $excelFile ---"

Order-Columns -data $removedAccounts   -Order $desiredOrder |
    Export-Excel -Path $excelFile -WorksheetName "Removed" -AutoSize -Title "Removed Users"

Order-Columns -data $newAccounts       -Order $desiredOrder |
    Export-Excel -Path $excelFile -WorksheetName "Added"   -AutoSize -Title "New Accounts"   -Append

Order-Columns -data $rehireAccounts    -Order $desiredOrder |
    Export-Excel -Path $excelFile -WorksheetName "Rehires" -AutoSize -Title "Rehired/Existing Accounts" -Append

if ($global:RehireChangeLog.Count) {
    $global:RehireChangeLog |
        Export-Excel -Path $excelFile -WorksheetName "RehireChanges" -AutoSize -Title "Before vs After" -Append
} else {
    @() | Export-Excel -Path $excelFile -WorksheetName "RehireChanges" -Title "No Rehire Changes" -Append
}

Write-Host "`nExport done => $excelFile"
Write-Host "All done. Real AD changes applied? $WillChangeAD (false = simulation only)."
