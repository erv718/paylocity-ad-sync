param(
    [string]$SyncReportPath = "C:\Reports\SyncReport_*.xlsx"
)

Import-Module ActiveDirectory -ErrorAction Stop
Import-Module ImportExcel   -ErrorAction Stop

# Pick the newest report
$report = Get-ChildItem $SyncReportPath |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (-not $report) {
    Write-Error "No SyncReport found matching $SyncReportPath"
    return
}

Write-Host "Using report:" $report.FullName
Write-Host ""

# 1) Load sheets, skipping the title row
$removed = Import-Excel -Path $report.FullName -WorksheetName "Removed" -StartRow 2
$added   = Import-Excel -Path $report.FullName -WorksheetName "Added"   -StartRow 2
$rehire  = Import-Excel -Path $report.FullName -WorksheetName "Rehires" -StartRow 2

# 2) Show actual headers and first few rows
Write-Host "=== 'Removed' sheet columns ==="
$removed[0].PSObject.Properties.Name | ForEach-Object { Write-Host " - $_" }

Write-Host "`n=== First 5 rows of 'Removed' ==="
$removed | Select-Object -First 5 | Format-List

# 3) Off-board debug
Write-Host "`n=== OFFBOARD DEBUG ==="
foreach ($r in $removed) {
    $id    = $r.Role
    $fname = $r.'First Name'
    $lname = $r.'Last Name'
    Write-Host "Checking ID '$id' ($fname $lname)..."

    # skip if no ID
    if (-not $id) {
        Write-Warning "  -> No Role value, skipping"
        continue
    }

    $matches = Get-ADUser `
        -Filter "(employeeID -eq '$id') -or (employeeNumber -eq '$id')" `
        -Properties SamAccountName,DistinguishedName `
        -ErrorAction SilentlyContinue

    switch ($matches.Count) {
        0 { Write-Host "  -> No AD user found" }
        1 { Write-Host ("  -> One match: SamAccountName={0}, DN={1}" -f $matches.SamAccountName, $matches.DistinguishedName) }
        default {
            Write-Warning ("  -> Multiple matches: {0}" -f $matches.Count)
            $matches | Format-Table SamAccountName, DistinguishedName -AutoSize
        }
    }
}

# 4) New-user param dump
Write-Host "`n=== NEW-USER PARAMS DEBUG ==="
foreach ($a in $added) {
    $id    = $a.Role
    if (-not $id) {
        Write-Warning "Skipped new-user row with no Role"
        continue
    }
    $first = if ($a.'Preferred First Name') { $a.'Preferred First Name' } else { $a.'First Name' }
    $last  = $a.'Last Name'
    $sam   = ("{0}.{1}" -f $first, $last).Replace(' ', '').ToLower()
    $email = "$sam@corp.example.com"

    $params = @{
        Name              = "$first $last"
        SamAccountName    = $sam
        UserPrincipalName = "$sam@ad.corp.example.com"
        GivenName         = $first
        Surname           = $last
        EmailAddress      = $email
        MobilePhone       = $a.'Mobile Phone'
        Description       = "$($a.'Location  Code') - $($a.'Job Title')"
        AccountPassword   = '<SecureString>'
        Enabled           = $true
        Path              = "OU=Users,DC=ad,DC=corp.example,DC=com"
        # OtherAttributes = @{
        #     employeeID = $id
        #     mail       = $email  # <- DO NOT use if you use -EmailAddress
        # }
    }

    Write-Host ("Would run New-ADUser for ID {0} ({1} {2}):" -f $id, $first, $last)
    $params.GetEnumerator() | Format-Table Name,Value -AutoSize
    Write-Host "Note: remove OtherAttributes['mail'] if you also use -EmailAddress."
}

# 5) Rehire-update debug
Write-Host "`n=== REHIRE-UPDATE DEBUG ==="
foreach ($a in $rehire) {
    $id    = $a.Role
    if (-not $id) {
        Write-Warning "Skipped rehire row with no Role"
        continue
    }
    $fname = $a.'First Name'
    $lname = $a.'Last Name'
    Write-Host "Checking ID '$id' ($fname $lname)..."

    $u = Get-ADUser `
        -Filter "(employeeID -eq '$id') -or (employeeNumber -eq '$id')" `
        -Properties givenName,sn,title,department,mail `
        -ErrorAction SilentlyContinue

    if ($u.Count -ne 1) {
        Write-Warning ("  -> Expected 1 match, got {0}" -f $u.Count)
        continue
    }
    $u = $u[0]

    Write-Host "  Existing AD values:"
    [pscustomobject]@{
        givenName  = $u.givenName
        sn         = $u.sn
        title      = $u.title
        department = $u.department
        mail       = $u.mail
    } | Format-Table -AutoSize

    $replaceHash = @{
        givenName  = $fname
        sn         = $lname         # must use LDAP 'sn'
        title      = $a.'Job Title'
        department = $a.'Department Name'
        mail       = $a.Email
    }
    Write-Host "  Would call Set-ADUser -Replace with:"
    $replaceHash.GetEnumerator() | Format-Table Name,Value -AutoSize
    Write-Host "Note: use LDAP attribute 'sn', not 'surname'."
}

Write-Host "`nDebug complete." -ForegroundColor Green
