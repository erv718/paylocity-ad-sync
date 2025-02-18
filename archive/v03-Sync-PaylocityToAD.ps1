<#
    Example script that:
    1) Compares previous & current CSV
    2) Identifies removed vs. added records
    3) Splits added accounts into New Accounts vs. Rehires/Existing
    4) Logs results in an Excel file (with multiple tabs, with ordered columns)
    5) Optionally updates AD if -NoWhatIf is provided
    6) Normalizes Role numbers (adds a leading zero for 5-digit values)
    7) Skips users with Department Name = "Personal Training", "Maintenance", or "Front Desk"
    8) Handles missing/blank headers by re-importing with a predefined header list
    9) Trims column headers and values to handle spacing issues
    10) Skips any duplicate header rows (not just the first row)
#>

[CmdletBinding()]
param(
    [string]$PreviousCsv = "C:\Reports\Employee Data 2025-02-20.csv",
    [string]$CurrentCsv  = "C:\Reports\Employee Data 2025-02-21.csv",
    [switch]$NoWhatIf
)

# Import required modules
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

# Determine if AD changes will be made
$WillChangeAD = $NoWhatIf.IsPresent
Write-Host "WillChangeAD = $WillChangeAD (false => simulation only)"

# --- Helper: Normalize Records ---
function Normalize-Records {
    param(
         [Parameter(Mandatory = $true)]
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

# --- Enhanced Import Function with Duplicate Header Filtering ---
function Import-CleanCsv {
    param(
        [string]$Path,
        [string]$Delimiter = ",",
        # Predefined header list matching sample data columns
        [string[]]$ExpectedHeaders = @(
            "Role","Email","Personal Email","First Name","Preferred First Name","Last Name",
            "Mobile Phone","Location  Code","Location Description","Department Code",
            "Department Name","Job Title","Salary or Hourly","Original Hire Date","DOB",
            "Type","Training Phase","address1","address2","city","state","zip",
            "rehiredate","status","Termination Date","isexempt","issalaried","ishourly"
        )
    )
    # 1) Import CSV normally
    $data = Import-Csv -Path $Path -Delimiter $Delimiter -ErrorAction SilentlyContinue
    $data = Normalize-Records -records $data

    # 2) Check if all expected headers are present; if not, re-import with specified headers.
    if ($data.Count -gt 0) {
        $headersPresent = $data[0].PSObject.Properties.Name
        $missingHeaders = $ExpectedHeaders | Where-Object { $_ -notin $headersPresent }
        if ($missingHeaders.Count -gt 0) {
            Write-Host "Missing headers: $($missingHeaders -join ', '). Re-importing with specified headers."
            $data = Import-Csv -Path $Path -Delimiter $Delimiter -Header $ExpectedHeaders -ErrorAction SilentlyContinue
            $data = Normalize-Records -records $data
        }
    }

    # 3) Remove duplicate header rows.
    # Build an array of expected headers in lower-case.
    $expectedHeadersLower = $ExpectedHeaders | ForEach-Object { $_.Trim().ToLower() }
    if ($data.Count -gt 0) {
        $data = $data | Where-Object {
            # Build a dictionary for current row with lower-case keys and values.
            $dict = @{}
            foreach ($prop in $_.PSObject.Properties) {
                $keyLower = $prop.Name.Trim().ToLower()
                $dict[$keyLower] = ($prop.Value -as [string]).Trim().ToLower()
            }
            $isHeaderRow = $true
            foreach ($header in $expectedHeadersLower) {
                if (-not $dict.ContainsKey($header) -or $dict[$header] -ne $header) {
                    $isHeaderRow = $false
                    break
                }
            }
            -not $isHeaderRow
        }
    }

    return $data
}

# --- Helper: Order Columns for Excel Export ---
function Order-Columns {
    param(
       [Parameter(Mandatory = $true)]
       [array]$data,
       [Parameter(Mandatory = $true)]
       [string[]]$Order
    )
    if ($data.Count -gt 0) {
       $existing = $data[0].PSObject.Properties.Name
       $columns = $Order | Where-Object { $existing -contains $_ }
       return $data | Select-Object -Property $columns
    }
    else {
       return $data
    }
}

# Define desired column order (adjust as needed)
$desiredOrder = @(
    "Role","Preferred First Name","First Name","Last Name",
    "Email","Personal Email","Mobile Phone",
    "Location  Code","Location Description","Department Code",
    "Department Name","Job Title","address1","address2",
    "city","state","zip","rehiredate","status","Termination Date",
    "isexempt","issalaried","ishourly"
)

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

# 6) Process REMOVED (Offboarding)
if ($removed.Count -gt 0) {
    Write-Host "`n--- REMOVED ---"
    foreach ($r in $removed) {
        $id = $r.Role
        $fn = $r.'First Name'
        $ln = $r.'Last Name'
        Write-Host " - [$id] $fn $ln => Would disable/move in AD"
        if ($WillChangeAD) {
            # Lookup AD user using employeeID, employeeNumber, or mail.
            $filter = "(employeeID -eq '$($id)') -or (employeeNumber -eq '$($id)') -or (mail -eq '$($r.Mail)')"
            $adUser = Get-ADUser -Filter $filter -Properties SamAccountName -ErrorAction SilentlyContinue
            if ($adUser) {
                Disable-ADAccount -Identity $adUser.SamAccountName
                Move-ADObject -Identity $adUser.DistinguishedName -TargetPath "OU=Users,OU=Disabled,OU=[Company],DC=ad,DC=corp.example,DC=com"
                # Remove from groups except Domain Users
                $groups = Get-ADUser $adUser.SamAccountName -Properties MemberOf | Select-Object -ExpandProperty memberOf
                foreach ($dn in $groups) {
                    if (-not ($dn -like "*Domain Users*")) {
                        Remove-ADGroupMember -Identity $dn -Members $adUser.SamAccountName -Confirm:$false
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

# 7) Process ADDED (Onboarding / Rehire)
if ($added.Count -gt 0) {
    Write-Host "`n--- ADDED (Processing New vs. Rehire) ---"
    foreach ($a in $added) {
        $id   = $a.Role
        $fn   = $a.'First Name'
        $ln   = $a.'Last Name'
        $mail = $a.Mail
        Write-Host " - [$id] $fn $ln => Checking if rehire vs new"
        # Build filter based on employeeID and mail.
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
                # Use Preferred First Name if exists; otherwise, use First Name.
                $firstName = $a.'Preferred First Name'
                if ([string]::IsNullOrWhiteSpace($firstName)) {
                    $firstName = $a.'First Name'
                }
                $lastName = $a.'Last Name'
                # Construct email: remove spaces, lower-case.
                $emailLocal = ($firstName + "." + $lastName) -replace "\s", ""
                $email = "$emailLocal@corp.example.com"
                # Construct Description: "Location Code - Job Title"
                $description = "$($a.'Location  Code') - $($a.'Job Title')"
                # Build a simple samAccountName; adjust as needed.
                $samAccountName = $emailLocal.ToLower()
                # Build hash table of attributes; note: Personal Email goes into extensionAttribute2.
                $userAttrs = @{
                    employeeID          = $id
                    mail                = $email
                    title               = $a.'Job Title'
                    department          = $a.'Department Name'
                    mobile              = $a.'Mobile Phone'
                    description         = $description
                    extensionAttribute2 = $a.'Personal Email'
                }
                Write-Host "   Creating new AD user: $samAccountName ($firstName $lastName)"
                New-ADUser -Name "$firstName $lastName" `
                           -SamAccountName $samAccountName `
                           -UserPrincipalName "$samAccountName@ad.corp.example.com" `
                           -GivenName $firstName `
                           -Surname $lastName `
                           -EmailAddress $email `
                           -MobilePhone $a.'Mobile Phone' `
                           -Description $description `
                           -AccountPassword (ConvertTo-SecureString "P@ssw0rd123" -AsPlainText -Force) `
                           -Enabled $true `
                           -Path "OU=Users,DC=ad,DC=corp.example,DC=com" `
                           -OtherAttributes $userAttrs
                Write-Host "   New AD user created in default OU."
                # Add to security group "sec00us-googleuser-sec"
                Add-ADGroupMember -Identity "sec00us-googleuser-sec" -Members $samAccountName -ErrorAction SilentlyContinue
                Write-Host "   Added to security group 'sec00us-googleuser-sec'."
            }
        }
    }
}
else {
    Write-Host "No added users."
}

# 8) Export to Excel with multiple tabs (with ordered columns)
$timeStamp = (Get-Date -Format "yyyy-MM-dd_HH-mm")
$excelFile = "C:\Reports\SyncReport_$timeStamp.xlsx"
Write-Host "`n--- Exporting to $excelFile with multiple tabs ---"

# Removed
if ($removed) {
    $removedOrdered = Order-Columns -data $removed -Order $desiredOrder
    $removedOrdered | Export-Excel -Path $excelFile -WorksheetName "Removed" -AutoSize -Title "Removed Users"
}
else {
    $nullArr = New-Object System.Collections.ArrayList
    $nullArr | Export-Excel -Path $excelFile -WorksheetName "Removed" -AutoSize -Title "Removed Users"
}

# New Accounts (Added - new)
if ($newAccounts) {
    $newOrdered = Order-Columns -data $newAccounts -Order $desiredOrder
    $newOrdered | Export-Excel -Path $excelFile -WorksheetName "Added" -AutoSize -Title "New Accounts" -Append
}
else {
    $nullArr = New-Object System.Collections.ArrayList
    $nullArr | Export-Excel -Path $excelFile -WorksheetName "Added" -AutoSize -Title "New Accounts" -Append
}

# Rehires
if ($rehireAccounts) {
    $rehireOrdered = Order-Columns -data $rehireAccounts -Order $desiredOrder
    $rehireOrdered | Export-Excel -Path $excelFile -WorksheetName "Rehires" -AutoSize -Title "Rehired/Existing Accounts" -Append
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
Write-Host "`nAll done. Real AD changes applied? $WillChangeAD (false = simulation only)."
