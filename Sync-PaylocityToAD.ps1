<#
.SYNOPSIS
    Sync AD users from Paylocity_Data_Latest.xlsx

.DESCRIPTION
    • Reads only Paylocity_Data_Latest.xlsx (skips title/footer rows)
    • Splits by status (A/T) into Removed/New/Rehire via AD lookups
    • Logs to SyncReport_*.xlsx with multiple tabs & ordered columns
    • Optional AD changes with –Apply (alias –NoWhatIf)
    • Normalizes 5-digit Role numbers (adds leading zero) and checks both forms
    • Skips Personal Training, Maintenance, Front Desk
    • Handles missing/blank headers; trims; removes duplicates
    • Drops final "Report Total Records: N" footer
    • Default password = FirstInitial + LastInitial + "[company]" + last 4 of ID
    • Places new users under Corporate vs. Clubs OUs based on Location Description
    • Stores Personal Email in extensionAttribute2
    • Appends suffix (01,02,…) if SamAccountName already exists
#>

[CmdletBinding()]
param(
    [string]$Report = "C:\Reports\Paylocity_Data_Latest.xlsx",
    [Alias("NoWhatIf")][switch]$Apply,
    [Alias("Mod")][switch]$Modify
)

# 0) Modules & basic checks
Import-Module ActiveDirectory -ErrorAction Stop

# Ensure ImportExcel is available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing from PSGallery…" -ForegroundColor Yellow
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    }
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel -ErrorAction Stop

Write-Host "Report => $Report"
if (-not (Test-Path $Report)) {
    Write-Error "Report not found: $Report"
    return
}

$WillChangeAD = $Apply.IsPresent
$adText = if ($WillChangeAD) { 'YES' } else { 'NO' }
Write-Host ("WillChangeAD = {0} (false => simulation only)" -f $WillChangeAD)

# Global rehire-change log
$global:RehireChangeLog = @()
# Global modify-report log for updated active accounts
$global:ModifyReport = @()
# Global error log
$global:ErrorLog = @()
# Global duplicate log
$global:DuplicateLog = @()

# 1) Normalize trim helper
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
    # Skip title row, treat row2 as headers
    $data = Import-Excel -Path $Path -StartRow 2 -ErrorAction SilentlyContinue
    $data = Normalize-Records -records $data
    if (-not $data.Count) { return @() }

    # If headers missing, re-import with predefined
    $present = $data[0].PSObject.Properties.Name
    $missing = $ExpectedHeaders | Where-Object { $_ -notin $present }
    if ($missing) {
        Write-Host "Missing headers: $($missing -join ', '). Re-importing with predefined headers."
        $data = Import-Excel -Path $Path -StartRow 2 -Header $ExpectedHeaders -ErrorAction SilentlyContinue
        $data = Normalize-Records -records $data
    }

    # Drop duplicate-header rows
    $expLow = $ExpectedHeaders | % { $_.Trim().ToLower() }
    $data = $data | Where-Object {
        $map = @{}
        foreach ($p in $_.PSObject.Properties) {
            $k = $p.Name.Trim().ToLower()
            $v = ($p.Value -as [string]) -replace '^\s+|\s+$',''
            $map[$k] = $v.ToLower()
        }
        $isHdr = $true
        foreach ($h in $expLow) {
            if (-not $map.ContainsKey($h) -or $map[$h] -ne $h) {
                $isHdr = $false; break
            }
        }
        -not $isHdr
    }

    # Drop "Report Total Records: N" footer
    $script:ReportTotalCount = $null
    $data = $data | Where-Object {
        if ($_.Role -match '^Report\s*Total\s*Records:\s*(\d+)$') {
            $script:ReportTotalCount = [int]$Matches[1]
            $false
        } else { $true }
    }

    # Drop stray single-cell rows
    $data = $data | Where-Object {
        ($_.PSObject.Properties.Value | Where-Object { $_ -notin @($null,'') }).Count -gt 1
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

# Skip these departments
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
# Detect duplicate Roles in input
$data | Group-Object -Property Role | Where-Object { $_.Count -gt 1 } | ForEach-Object {
    $global:DuplicateLog += [PSCustomObject]@{
        Role  = $_.Name
        Count = $_.Count
    }
}

# 5) Classify
$removedAccounts = @(); $newAccounts = @(); $rehireAccounts = @(); $modifyAccounts = @()
foreach ($r in $data) {
    if ($skipDepartments -contains $r.'Department Name') { continue }
    if (-not $r.Role) { continue }
    $rawID = $r.Role.Trim()
    if ($rawID -match '^\d{5}$') { $r.Role = "0$rawID" }

    if (-not $r.'First Name' -or -not $r.'Last Name') { continue }
    switch ($r.status) {
        'T' {
            $removedAccounts += $r
        }
        'A' {
            $f = "((employeeID -eq '$($r.Role)') -or (employeeNumber -eq '$($r.Role)')" +
                 " -or (employeeID -eq '$rawID') -or (employeeNumber -eq '$rawID'))"
            $u = Get-ADUser -Filter $f -Properties Enabled -ErrorAction SilentlyContinue
            if (-not $u)             { $newAccounts    += $r }
            elseif (-not $u.Enabled) { $rehireAccounts += $r }
            else                     { $modifyAccounts += $r }
        }
        default {
            Write-Host "[-] Unknown status '$($r.status)' for Role $($r.Role). Skipping."
            $global:ErrorLog += [PSCustomObject]@{ Type = 'UnknownStatus'; Role = $r.Role; Status = $r.status }
        }
    }
}

# 6) SUMMARY
Write-Host "`n=== SUMMARY ==="
Write-Host ("Terminated: {0}" -f $removedAccounts.Count)
Write-Host ("New Accounts: {0} | Rehires: {1}" -f $newAccounts.Count, $rehireAccounts.Count)
Write-Host ("AD changes? {0}" -f $adText)

# 7) OFFBOARDING
if ($removedAccounts.Count) {
    Write-Host "`n--- OFFBOARDING ---"
    foreach ($r in $removedAccounts) {
        $id = $r.Role; $fn = $r.'First Name'; $ln = $r.'Last Name'
        Write-Host " - [$id] $fn $ln => disable/move in AD"
        $all = Get-ADUser -Filter "(employeeID -eq '$id') -or (employeeNumber -eq '$id')" `
              -Properties SamAccountName,DistinguishedName,Enabled -ErrorAction SilentlyContinue
        $act = $all | Where-Object { $_.Enabled }
        switch ($act.Count) {
            1 {
                $u = $act[0]
                if ($WillChangeAD) {
                    Disable-ADAccount -Identity $u.SamAccountName
                    Move-ADObject -Identity $u.DistinguishedName `
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
                Write-Warning "No active AD user found for Role $id - skipping off-boarding"
                $global:ErrorLog += [PSCustomObject]@{ Type = 'OffboardMissing'; Role = $id }
            }
            default {
                Write-Warning "Multiple active matches ($($act.Count)) for Role $id - skipping off-boarding"
                $global:ErrorLog += [PSCustomObject]@{ Type = 'OffboardMultiple'; Role = $id; Count = $act.Count }
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
        $desc  = "$($r.'Location  Code') - $($r.'Job Title')"
        Write-Host " - [$id] $fn $ln => re-enable & update"
        # Re-enable filter: include both padded and raw IDs and close parentheses
        $rawID2 = $r.Role.Trim()
        if ($rawID2 -match '^0\d{5}$') { $rawID2 = $rawID2.Substring(1) }
        $f = "((employeeID -eq '$($r.Role)') -or (employeeNumber -eq '$($r.Role)') -or (employeeID -eq '$rawID2') -or (employeeNumber -eq '$rawID2'))"
        $usr = Get-ADUser -Filter $f -Properties Enabled,givenName,sn,title,department,mail `
               -ErrorAction SilentlyContinue
        if (-not $usr) {
            Write-Warning "AD user not found for rehire."
            $global:ErrorLog += [PSCustomObject]@{ Type = 'RehireMissing'; Role = $id }
            continue
        }
        if ($WillChangeAD) {
            $before = Get-ADUser -Identity $usr.SamAccountName -Properties givenName,sn,title,department,mail,company,departmentNumber,employeeID,mobile,otherMailbox,description,displayName
            $newH = @{
                givenName        = $fn
                sn               = $ln
                title            = $r.'Job Title'
                department       = $r.'Department Name'
                company          = '[Company]'
                departmentNumber = $r.'Location  Code'
                employeeID       = $id
                mobile           = $r.'Mobile Phone'
                otherMailbox     = $r.'Personal Email'
                description      = $desc
                displayName      = "$fn $ln"
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
            Set-ADUser -Identity $usr.SamAccountName -Replace $newH
            if ($r.'Location Description' -eq 'Corporate') {
                Add-ADGroupMember -Identity "sec00us-GoogleLicenseEnterpriseStandard-sec" -Members $usr.SamAccountName -ErrorAction Stop
                Write-Host "   Added to sec00us-GoogleLicenseEnterpriseStandard-sec."
            } else {
                Add-ADGroupMember -Identity "sec00us-GoogleLicenseFrontlineVault-sec" -Members $usr.SamAccountName -ErrorAction Stop
                Write-Host "   Added to sec00us-GoogleLicenseFrontlineVault-sec."
            }
            Write-Host "   Re-enabled & updated."
        } else {
            Write-Host "   [Simulation] Would re-enable & update."
        }
    }
} else {
    Write-Host "No rehires."
}

# Insert modify block here
if ($Modify.IsPresent) {
    if ($modifyAccounts.Count) {
        Write-Host "`n--- MODIFY EXISTING ACTIVE ACCOUNTS ---"
        foreach ($r in $modifyAccounts) {
            $id = $r.Role
            # Use Preferred First Name if available
            $fn = if ($r.'Preferred First Name') { $r.'Preferred First Name' } else { $r.'First Name' }
            $ln = $r.'Last Name'
            $desc = "$($r.'Location  Code') - $($r.'Job Title')"
            Write-Host " - [$id] $fn $ln => modify attributes"
            $rawID2 = $r.Role.Trim()
            if ($rawID2 -match '^0\d{5}$') { $rawID2 = $rawID2.Substring(1) }
            $f2 = "((employeeID -eq '$($r.Role)') -or (employeeNumber -eq '$($r.Role)') -or (employeeID -eq '$rawID2') -or (employeeNumber -eq '$rawID2'))"
            $usr2 = Get-ADUser -Filter $f2 -Properties givenName,sn,title,department,mail,company,departmentNumber,employeeID,mobile,otherMailbox,description,displayName -ErrorAction SilentlyContinue
            if (-not $usr2) {
                Write-Warning "AD user not found for modify."
                $global:ErrorLog += [PSCustomObject]@{ Type = 'ModifyMissing'; Role = $id }
                continue
            }
            # Prepare updated attributes (exclude mail and otherMailbox)
            $newH = @{                
                givenName           = $fn
                sn                  = $ln
                title               = $r.'Job Title'
                department          = $r.'Department Name'
                company             = '[Company]'
                departmentNumber    = $r.'Location  Code'
                mobile              = $r.'Mobile Phone'
                extensionAttribute2 = $r.'Personal Email'
                description         = $desc
                displayName         = "$fn $ln"
            }
            # Always log new attribute values for report (even in simulation)
            $global:ModifyReport += [PSCustomObject]@{
                SamAccountName     = $usr2.SamAccountName
                FirstName          = $newH.givenName
                LastName           = $newH.sn
                DisplayName        = $newH.displayName
                Title              = $newH.title
                Department         = $newH.department
                Company            = $newH.company
                DepartmentNumber   = $newH.departmentNumber
                EmployeeID         = "'" + $newH.employeeID
                Mobile             = $newH.mobile
                ExtensionAttribute2= $newH.extensionAttribute2
                Description        = $newH.description
            }
            if ($WillChangeAD) {
                Set-ADUser -Identity $usr2.SamAccountName -Replace $newH
                if ($r.'Location Description' -eq 'Corporate') {
                    Add-ADGroupMember -Identity "sec00us-GoogleLicenseEnterpriseStandard-sec" -Members $usr2.SamAccountName -ErrorAction Stop
                    Write-Host "   Added to sec00us-GoogleLicenseEnterpriseStandard-sec."
                } else {
                    Add-ADGroupMember -Identity "sec00us-GoogleLicenseFrontlineVault-sec" -Members $usr2.SamAccountName -ErrorAction Stop
                    Write-Host "   Added to sec00us-GoogleLicenseFrontlineVault-sec."
                }
                Write-Host "   Modified attributes."
            } else {
                Write-Host "   [Simulation] Would modify attributes."
            }
        }
    } else {
        Write-Host "No existing active accounts to modify."
    }
}

# 9) ONBOARD NEW ACCOUNTS
if ($newAccounts.Count) {
    Write-Host "`n--- ONBOARDING NEW ACCOUNTS ---"
    foreach ($r in $newAccounts) {
        $id = $r.Role
        $fn = if ($r.'Preferred First Name') { $r.'Preferred First Name' } else { $r.'First Name' }
        $ln = $r.'Last Name'
        # build UPN and pre-Windows2000 SamAccountName
        $baseUpn = ("{0}.{1}" -f $fn,$ln).ToLower() -replace '[^a-z0-9\.]',''
        if ($baseUpn.Length -gt 20) { $baseUpn = $baseUpn.Substring(0,20) }
        $upn = $baseUpn
        $basePreSam = ($fn.Substring(0,1) + $ln) -replace '[^a-zA-Z0-9]',''
        if ($basePreSam.Length -gt 20) { $basePreSam = $basePreSam.Substring(0,20) }
        $preSam = $basePreSam
        $psSuffix = 1
        while (Get-ADUser -Filter "SamAccountName -eq '$preSam'" -ErrorAction SilentlyContinue) {
            $preSam = $basePreSam + ('{0:00}' -f $psSuffix)
            $psSuffix++
            if ($psSuffix -gt 99) { break }
        }
        $email = "$upn@corp.example.com"
        $desc  = "$($r.'Location  Code') - $($r.'Job Title')"
        # Generate last four of Role (or full ID if shorter)
        $start = [Math]::Max(0, $id.Length - 4)
        $last4 = $id.Substring($start)
        $pw    = "{0}{1}[company]{2}" -f $fn[0], $ln[0], $last4

        # choose OU by Location Description
        if ($r.'Location Description' -eq 'Corporate') {
            $ou = "OU=Users,OU=Corporate,OU=[Company],DC=SANDBOX,DC=local"
        } else {
            $ou = "OU=Users,OU=Gym,OU=[Company],DC=SANDBOX,DC=local"
        }

        Write-Host " - [$id] UPN:$upn  sAMAccountName:$preSam ($fn $ln) => create in $ou"
        if ($WillChangeAD) {
            try {
                New-ADUser `
                  -Name             "$fn $ln" `
                  -DisplayName      "$fn $ln" `
                  -GivenName        $fn `
                  -Surname          $ln `
                  -SamAccountName   $preSam `
                  -UserPrincipalName "$upn@ad.corp.example.com" `
                  -EmailAddress     $email `
                  -MobilePhone      $r.'Mobile Phone' `
                  -Description      $desc `
                  -Company          "[Company]" `
                  -Title            $r.'Job Title' `
                  -Department       $r.'Department Name' `
                  -AccountPassword  (ConvertTo-SecureString $pw -AsPlainText -Force) `
                  -Enabled          $true `
                  -Path             $ou `
                  -OtherAttributes  @{ 
                      employeeID         = $id
                      departmentNumber   = $r.'Location  Code'
                      extensionAttribute2 = $r.'Personal Email'
                  }
                  Add-ADGroupMember -Identity "sec00us-googleuser-sec" -Members $preSam -ErrorAction Stop
                  Write-Host "   Created & added to sec00us-googleuser-sec."
                  if ($r.'Location Description' -eq 'Corporate') {
                      Add-ADGroupMember -Identity "sec00us-GoogleLicenseEnterpriseStandard-sec" -Members $preSam -ErrorAction Stop
                      Write-Host "   Added to sec00us-GoogleLicenseEnterpriseStandard-sec."
                  } else {
                      Add-ADGroupMember -Identity "sec00us-GoogleLicenseFrontlineVault-sec" -Members $preSam -ErrorAction Stop
                      Write-Host "   Added to sec00us-GoogleLicenseFrontlineVault-sec."
                  }
              } catch {
                  Write-Warning "   ✗ Failed to create ${preSam}: $($_.Exception.Message)"
              }
          } else {
              Write-Host "   [Simulation] Would create $preSam with password '$pw'."
          }
      }
  } else {
      Write-Host "No new accounts."
  }

  # 10) EXPORT
  $ts  = Get-Date -Format "yyyy-MM-dd_HH-mm"
  $out = "C:\Reports\SyncReport_$ts.xlsx"
  Write-Host "`n--- Exporting to $out ---"

  # Removed
  Order-Columns -data $removedAccounts -Order $desiredOrder |
    Export-Excel -Path $out -WorksheetName "Removed" -AutoSize -Title "Removed Users"
  # Added
  Order-Columns -data $newAccounts -Order $desiredOrder |
    Export-Excel -Path $out -WorksheetName "Added" -AutoSize -Title "New Accounts" -Append
  # Rehires
  Order-Columns -data $rehireAccounts -Order $desiredOrder |
    Export-Excel -Path $out -WorksheetName "Rehires" -AutoSize -Title "Rehired/Existing" -Append
  # RehireChanges
  if ($global:RehireChangeLog.Count) {
      $global:RehireChangeLog |
        Export-Excel -Path $out -WorksheetName "RehireChanges" -AutoSize -Title "Before vs After" -Append
  } else {
      @() |
        Export-Excel -Path $out -WorksheetName "RehireChanges" -Title "No Rehire Changes" -Append
  }
  # Modified changes log
  if ($global:ModifyReport.Count) {
      $global:ModifyReport |
        Export-Excel -Path $out -WorksheetName "Modified" -AutoSize -Title "New Attributes" -Append
  } else {
      @() |
        Export-Excel -Path $out -WorksheetName "Modified" -Title "No Modified Accounts" -Append
  }

  # Errors log
  if ($global:ErrorLog.Count) {
      $global:ErrorLog |
        Export-Excel -Path $out -WorksheetName "Errors" -AutoSize -Title "Error Log" -Append
  } else {
      @() |
        Export-Excel -Path $out -WorksheetName "Errors" -Title "No Errors" -Append
  }
  # Duplicates log
  if ($global:DuplicateLog.Count) {
      $global:DuplicateLog |
        Export-Excel -Path $out -WorksheetName "Duplicates" -AutoSize -Title "Duplicate Roles" -Append
  } else {
      @() |
        Export-Excel -Path $out -WorksheetName "Duplicates" -Title "No Duplicates" -Append
  }

  Write-Host "`nDone. Real AD changes applied? $WillChangeAD (false = simulation only)."