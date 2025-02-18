<#
    Example script that:
    1) Compares previous & current CSV
    2) Identifies removed vs. added records
    3) Splits added accounts into New Accounts vs. Rehires/Existing
    4) Logs results in an Excel file (with multiple tabs)
    5) Optionally updates AD if -NoWhatIf is provided
    6) Normalizes Role numbers (adds a leading zero for 5-digit values)
    7) Skips users with Department Name = "Personal Training", "Maintenance", or "Front Desk"
    8) Handles missing/blank headers by re-importing with a predefined header list
    9) Trims column headers and values to handle spacing issues
    10) Skips a duplicate header row if detected
#>

[CmdletBinding()]
param(
    [string]$PreviousCsv = "C:\Reports\Employee Data 2025-02-19.csv",
    [string]$CurrentCsv  = "C:\Reports\Employee Data 2025-02-20.csv",
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

# --- Helper: Normalize Records ---
function Normalize-Records {
    param(
         [Parameter(Mandatory=$true)]
         [array]$records
    )
    return $records | ForEach-Object {
         $newProps = @{}
         foreach ($prop in $_.PSObject.Properties) {
             $newKey   = $prop.Name.Trim()        # Trim header name
             $newValue = $prop.Value
             if ($newValue -and ($newValue -is [string])) {
                  $newValue = $newValue.Trim()    # Trim string values
             }
             $newProps[$newKey] = $newValue
         }
         [PSCustomObject]$newProps
    }
}

# --- Enhanced Import Function ---
function Import-CleanCsv {
    param(
        [string]$Path,
        [string]$Delimiter = ",",
        # Below is the header list matching your sample data columns
        [string[]]$ExpectedHeaders = @(
            "Role","Email","Personal Email","First Name","Preferred First Name","Last Name",
            "Mobile Phone","Location  Code","Location Description","Department Code",
            "Department Name","Job Title","Salary or Hourly","Original Hire Date","DOB",
            "Type","Training Phase","address1","address2","city","state","zip",
            "rehiredate","status","Termination Date","isexempt","issalaried","ishourly"
        )
    )

    # 1) Import the CSV normally
    $data = Import-Csv -Path $Path -Delimiter $Delimiter -ErrorAction SilentlyContinue
    $data = Normalize-Records -records $data

    # 2) Check if all expected headers are present
    if ($data.Count -gt 0) {
        $headersPresent = $data[0].PSObject.Properties.Name
        $missingHeaders = $ExpectedHeaders | Where-Object { $_ -notin $headersPresent }
        if ($missingHeaders.Count -gt 0) {
            Write-Host "Missing headers: $($missingHeaders -join ', '). Re-importing with specified headers."
            $data = Import-Csv -Path $Path -Delimiter $Delimiter -Header $ExpectedHeaders -ErrorAction SilentlyContinue
            $data = Normalize-Records -records $data
        }
    }

    # 3) If the first record's Role equals "Role", skip it as a duplicate header
    if ($data.Count -gt 0 -and $data[0].Role -eq "Role") {
        Write-Host "Duplicate header row detected in file: $Path. Skipping first row."
        $data = $data | Select-Object -Skip 1
    }

    return $data
}

# 1) Import both CSVs using the enhanced function
$prev = Import-CleanCsv -Path $PreviousCsv
$curr = Import-CleanCsv -Path $CurrentCsv

# 2) Define departments to skip
$skipDepartments = @("Personal Training", "Maintenance", "Front Desk")

# 3) Build hash tables keyed by 'Role'
$prevHash = @{}
foreach ($u in $prev) {
    if ($skipDepartments -contains $u.'Department Name') { continue }
    if ($u.Role) {
        # Normalize Role: if 5 digits, add a leading 0
        if ($u.Role.Length -eq 5) { $u.Role = "0" + $u.Role }
        $prevHash[$u.Role] = $u
    }
}

$currHash = @{}
foreach ($u in $curr) {
    if ($skipDepartments -contains $u.'Department Name') { continue }
    if ($u.Role) {
        if ($u.Role.Length -eq 5) { $u.Role = "0" + $u.Role }
        $currHash[$u.Role] = $u
    }
}

# 4) Identify removed vs. added records
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

# 5) Split added into new accounts vs. rehires
$rehireAccounts = @()
$newAccounts    = @()

# We'll track rehire changes (before vs. after AD attribute updates)
$global:RehireChangeLog = @()

Write-Host "`n=== SUMMARY ==="
Write-Host ("Removed: {0} | Added: {1}" -f $removed.Count, $added.Count)
if ($WillChangeAD) {
    Write-Host "AD changes? YES (NoWhatIf set)."
} else {
    Write-Host "AD changes? NO (default => no changes)."
}

# 6) Process REMOVED
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
                # Remove from groups except Domain Users
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

# 7) Process ADDED
if ($added.Count -gt 0) {
    Write-Host "`n--- ADDED (Processing New vs. Rehire) ---"
    foreach ($a in $added) {
        $id   = $a.Role
        $fn   = $a.'First Name'
        $ln   = $a.'Last Name'
        $mail = $a.Mail
        Write-Host " - [$id] $fn $ln => Checking if rehire vs new"

        # Check AD
        $filtParts = @()
        if ($id)   { $filtParts += "(employeeID -eq '$($id)') -or (employeeNumber -eq '$($id)')" }
        if ($mail) { $filtParts += "(mail -eq '$($mail)')" }
        $fullFilt = if ($filtParts) { $filtParts -join ' -or ' } else { "*" }

        $adUser = Get-ADUser -Filter $fullFilt -Properties * -ErrorAction SilentlyContinue
        if ($adUser) {
            Write-Host "   Found in AD => rehire scenario => would re-enable & update attributes"
            $rehireAccounts += $a
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
                Enable-ADAccount -Identity $adUser.SamAccountName
                Set-ADUser -Identity $adUser.SamAccountName -Replace $newHash
                Write-Host "   Re-enabled & updated user."
            }
        }
        else {
            Write-Host "   Not in AD => new user => would create"
            $newAccounts += $a
            if ($WillChangeAD) {
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

# 8) Export to Excel with multiple tabs
$timeStamp = (Get-Date -Format "yyyy-MM-dd_HH-mm")
$excelFile = "C:\Reports\SyncReport_$timeStamp.xlsx"
Write-Host "`n--- Exporting to $excelFile with multiple tabs ---"

# Removed
if ($removed) {
    $removed | Export-Excel -Path $excelFile -WorksheetName "Removed" -AutoSize -Title "Removed Users"
}
else {
    $nullArr = New-Object System.Collections.ArrayList
    $nullArr | Export-Excel -Path $excelFile -WorksheetName "Removed" -AutoSize -Title "Removed Users"
}

# Added (New Accounts only)
if ($newAccounts) {
    $newAccounts | Export-Excel -Path $excelFile -WorksheetName "Added" -AutoSize -Title "New Accounts" -Append
}
else {
    $nullArr = New-Object System.Collections.ArrayList
    $nullArr | Export-Excel -Path $excelFile -WorksheetName "Added" -AutoSize -Title "New Accounts" -Append
}

# Rehires
if ($rehireAccounts) {
    $rehireAccounts | Export-Excel -Path $excelFile -WorksheetName "Rehires" -AutoSize -Title "Rehired/Existing Accounts" -Append
}
else {
    $nullArr = New-Object System.Collections.ArrayList
    $nullArr | Export-Excel -Path $excelFile -WorksheetName "Rehires" -AutoSize -Title "Rehired/Existing Accounts" -Append
}

# RehireChangeLog
if ($global:RehireChangeLog.Count -gt 0) {
    $global:RehireChangeLog | Export-Excel -Path $excelFile -WorksheetName "RehireChanges" -AutoSize -Title "Before vs After" -Append
}
else {
    $nullArr = New-Object System.Collections.ArrayList
    $nullArr | Export-Excel -Path $excelFile -WorksheetName "RehireChanges" -AutoSize -Title "No Rehire Changes" -Append
}

Write-Host "Export done => $excelFile"
Write-Host "`nAll done. Real AD changes? $WillChangeAD (false => no changes)."
