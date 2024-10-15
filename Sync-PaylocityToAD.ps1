<#
    Example script that:
    1) Compares previous & current CSV
    2) Identifies removed vs. added records
    3) Logs them in an Excel file (with multiple tabs)
    4) Optionally updates AD if -NoWhatIf is provided
#>

[CmdletBinding()]
param(
    [string]$PreviousCsv = "C:\Reports\Employee Data 2025-02-18.csv",
    [string]$CurrentCsv  = "C:\Reports\Employee Data 2025-02-19.csv",
    [switch]$NoWhatIf
)

Import-Module ActiveDirectory -ErrorAction Stop
Import-Module ImportExcel     -ErrorAction Stop

Write-Host "Previous CSV => $PreviousCsv"
Write-Host "Current  CSV => $CurrentCsv"

if (-not (Test-Path $PreviousCsv)) {
    Write-Error "Previous CSV not found: $PreviousCsv"
    return
}
if (-not (Test-Path $CurrentCsv)) {
    Write-Error "Current CSV not found: $CurrentCsv"
    return
}

# Determine if we actually do AD changes or not
$WillChangeAD = $NoWhatIf.IsPresent
Write-Host "WillChangeAD = $WillChangeAD (false => only logs, no AD changes)"

# Import CSV data
$prev = Import-Csv -Path $PreviousCsv
$curr = Import-Csv -Path $CurrentCsv

# Build hash tables keyed by 'Role' (or whatever unique column)
$prevHash = @{}
foreach ($u in $prev) {
    if ($u.Role) { $prevHash[$u.Role] = $u }
}
$currHash = @{}
foreach ($u in $curr) {
    if ($u.Role) { $currHash[$u.Role] = $u }
}

# Identify removed vs. added
$removed = @()
foreach ($key in $prevHash.Keys) {
    if (-not $currHash.ContainsKey($key)) {
        $removed += $prevHash[$key]
    }
}

$added = @()
foreach ($key in $currHash.Keys) {
    if (-not $prevHash.ContainsKey($key)) {
        $added += $currHash[$key]
    }
}

# We'll track rehire changes (before vs. after AD attribute changes)
$global:RehireChangeLog = @()

Write-Host "`n=== SUMMARY ==="
Write-Host ("Removed: {0} | Added: {1}" -f $removed.Count, $added.Count)
if ($WillChangeAD) {
    Write-Host "AD changes? YES (NoWhatIf set)."
} else {
    Write-Host "AD changes? NO (default => no changes)."
}

# 1) REMOVED
if ($removed.Count -gt 0) {
    Write-Host "`n--- REMOVED ---"
    foreach ($r in $removed) {
        $id = $r.Role
        $fn = $r.'First Name'
        $ln = $r.'Last Name'
        Write-Host " - [$id] $fn $ln => Would disable/move in AD"

        if ($WillChangeAD) {
            # Actual AD logic
            $filter = "(employeeID -eq '$($id)') -or (employeeNumber -eq '$($id)') -or (mail -eq '$($r.Mail)')"
            $adUser = Get-ADUser -Filter $filter -Properties SamAccountName -ErrorAction SilentlyContinue
            if ($adUser) {
                Disable-ADAccount -Identity $adUser.SamAccountName
                Move-ADObject -Identity $adUser.DistinguishedName -TargetPath "OU=Disabled,DC=ad,DC=corp.example,DC=com"
                # remove from groups except Domain Users
                $groups = Get-ADUser $adUser.SamAccountName -Properties MemberOf | Select-Object -ExpandProperty memberOf
                foreach ($dn in $groups) {
                    if (-not ($dn -like "*Domain Users*")) {
                        Remove-ADGroupMember $dn -Members $adUser.SamAccountName -Confirm:$false
                    }
                }
                Write-Host "   Disabled & moved to Disabled OU"
            }
            else {
                Write-Host "   [Warning] Not found in AD => cannot remove"
            }
        }
    }
}
else {
    Write-Host "No removed users."
}

# 2) ADDED
if ($added.Count -gt 0) {
    Write-Host "`n--- ADDED ---"
    foreach ($a in $added) {
        $id   = $a.Role
        $fn   = $a.'First Name'
        $ln   = $a.'Last Name'
        $mail = $a.Mail

        Write-Host " - [$id] $fn $ln => Checking if rehire vs new"

        # Build filter to see if user already in AD
        $filtParts = @()
        if ($id)   { $filtParts += "(employeeID -eq '$($id)') -or (employeeNumber -eq '$($id)')" }
        if ($mail) { $filtParts += "(mail -eq '$($mail)')" }
        $fullFilt = if ($filtParts) { $filtParts -join ' -or ' } else { "*" }

        $adUser = Get-ADUser -Filter $fullFilt -Properties * -ErrorAction SilentlyContinue
        if ($adUser) {
            # Rehire scenario
            Write-Host "   Found in AD => rehire scenario => would re-enable & update attributes"

            if ($WillChangeAD) {
                $attrsToCompare = 'givenName','surname','title','department','mail'
                $before = Get-ADUser -Identity $adUser.SamAccountName -Properties $attrsToCompare

                $newHash = @{
                    givenName  = $fn
                    surname    = $ln
                    title      = $a.'Job Title'
                    department = $a.'Department Name'
                    mail       = $mail
                }

                # Compare & store in $global:RehireChangeLog
                foreach ($attr in $attrsToCompare) {
                    $oldVal = $before.$attr
                    $newVal = $newHash[$attr]
                    if ($oldVal -ne $newVal) {
                        $global:RehireChangeLog += [PSCustomObject]@{
                            SamAccountName = $adUser.SamAccountName
                            Attribute      = $attr
                            Before         = $oldVal
                            After          = $newVal
                        }
                    }
                }

                # Re-enable & update
                Enable-ADAccount -Identity $adUser.SamAccountName
                Set-ADUser -Identity $adUser.SamAccountName -Replace $newHash
                Write-Host "   Re-enabled & updated user."
            }
        }
        else {
            # New user scenario
            Write-Host "   Not in AD => new user => would create"

            if ($WillChangeAD) {
                # Example SamAccountName
                $sam = ($fn.Substring(0,1) + $ln).ToLower() -replace "\s",""
                $upn = "$sam@ad.corp.example.com"

                New-ADUser -Name "$fn $ln" `
                           -SamAccountName $sam `
                           -UserPrincipalName $upn `
                           -GivenName $fn `
                           -Surname $ln `
                           -Enabled $true `
                           -AccountPassword (ConvertTo-SecureString "P@ssw0rd123" -AsPlainText -Force) `
                           -Path "OU=Users,OU=Corporate,OU=Locations,OU=[Company],DC=ad,DC=corp.example,DC=com" `
                           -OtherAttributes @{
                                employeeID = $id
                                mail       = $mail
                                title      = $a.'Job Title'
                                department = $a.'Department Name'
                           }
                Write-Host "   Created new AD user => $fn $ln"
            }
        }
    }
}
else {
    Write-Host "No added users."
}

# 3) Export to Single Excel (WITHOUT -NoNumberConversion)
$timeStamp = (Get-Date -Format "yyyy-MM-dd_HH-mm")
$excelFile = "C:\Reports\SyncReport_$timeStamp.xlsx"

Write-Host "`n--- Exporting to $excelFile with multiple tabs ---"

# Tab 1: Removed
if ($removed) {
    $removed | Export-Excel -Path $excelFile -WorksheetName "Removed" -AutoSize -Title "Removed Users"
}
else {
    $nullArr = New-Object System.Collections.ArrayList
    $nullArr | Export-Excel -Path $excelFile -WorksheetName "Removed" -AutoSize -Title "Removed Users"
}

# Tab 2: Added
if ($added) {
    $added | Export-Excel -Path $excelFile -WorksheetName "Added" -AutoSize -Title "Added Users" -Append
}
else {
    $nullArr = New-Object System.Collections.ArrayList
    $nullArr | Export-Excel -Path $excelFile -WorksheetName "Added" -AutoSize -Title "Added Users" -Append
}

# Tab 3: RehireChangeLog
if ($global:RehireChangeLog.Count -gt 0) {
    $global:RehireChangeLog | Export-Excel -Path $excelFile -WorksheetName "RehireChanges" -AutoSize -Title "Before vs After" -Append
}
else {
    $nullArr = New-Object System.Collections.ArrayList
    $nullArr | Export-Excel -Path $excelFile -WorksheetName "RehireChanges" -AutoSize -Title "No Rehire Changes" -Append
}

Write-Host "Export done => $excelFile"

Write-Host "`nAll done. Real AD changes? $WillChangeAD (false => no changes)."
