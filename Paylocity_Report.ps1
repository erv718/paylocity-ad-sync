# ──────────── Logging Setup ────────────
$LogFile = "C:\Reports\Paylocity\logs\PaylocityLog.txt"
$LogDir  = Split-Path $LogFile -Parent

# Ensure the log directory exists
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

# Heartbeat to prove the task fired
"===== TASK FIRED: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') =====" |
    Out-File -FilePath $LogFile -Append

# Start transcript so all output is captured; ignore errors if already running
Start-Transcript -Path $LogFile -Append -ErrorAction SilentlyContinue

# ──────────────── Configuration ────────────────
$SftpHost        = "your-sftp-host.server.transfer.amazonaws.com"
$SftpPort        = 22
$SftpUser        = "paylocity"
$PrivateKeyPath  = "C:\keys\paylocity_id_rsa"

$RemoteFilePath  = "/employee_export/EmployeeData.XLSX"
$DateStamp       = Get-Date -Format "yyyy-MM-dd"
$LocalDir        = "C:\Reports\Paylocity"
$ArchiveDir      = Join-Path $LocalDir "Archive"
$LocalFileName   = "Paylocity_Data_Latest.xlsx"
$LocalFilePath   = Join-Path $LocalDir $LocalFileName
$ArchiveFilePath = Join-Path $ArchiveDir "Paylocity_Data_$DateStamp.xlsx"

$emailFrom    = "reporting@corp.example.com"
$emailTo      = "techops@corp.example.com"
$emailCc      = "paylocityreport@corp.example.com"
$emailSubject = "Paylocity Report - $DateStamp"
$emailBody    = "Attached is the Paylocity export for $DateStamp."
$smtpServer   = "smtp-relay.example.com"

# ───────── Ensure folders exist ─────────
foreach ($d in @($LocalDir, $ArchiveDir)) {
    if (-not (Test-Path $d)) {
        New-Item -ItemType Directory -Path $d -Force | Out-Null
    }
}

# ──────── Load Posh-SSH Module ─────────
Import-Module Posh-SSH -ErrorAction Stop

# ───────── Master try/catch/finally ─────────
try {
    Write-Host "Connecting to $SftpHost…" -ForegroundColor Cyan

    # Key-based auth uses a dummy PSCredential
    $secureDummy = ConvertTo-SecureString "unused" -AsPlainText -Force
    $cred        = New-Object System.Management.Automation.PSCredential($SftpUser, $secureDummy)

    # Create the SFTP session
    $session = New-SFTPSession `
        -ComputerName $SftpHost `
        -Port         $SftpPort `
        -Credential   $cred `
        -KeyFile      $PrivateKeyPath `
        -ErrorAction  Stop

    # ─── Download ─────────────────────────────
    Write-Host "⬇Downloading remote file to $LocalDir…" -ForegroundColor Cyan
    Get-SFTPItem `
        -SessionId   $session.SessionId `
        -Path        $RemoteFilePath `
        -Destination $LocalDir `
        -Force

    # Clear ReadOnly and rename to “latest”
    $remoteName = Split-Path $RemoteFilePath -Leaf
    $downloaded = Join-Path $LocalDir $remoteName
    (Get-Item $downloaded).Attributes = "Normal"

    Write-Host "Overwriting latest file…" -ForegroundColor Cyan
    Move-Item `
        -Path        $downloaded `
        -Destination $LocalFilePath `
        -Force

    # ─── Archive ──────────────────────────────
    Write-Host "Archiving → $ArchiveFilePath" -ForegroundColor Yellow
    Copy-Item -Path $LocalFilePath -Destination $ArchiveFilePath -Force

    # ─── Email ────────────────────────────────
    Write-Host "Sending email…" -ForegroundColor Cyan
    $mailParams = @{
        From        = $emailFrom
        To          = $emailTo
        Cc          = $emailCc
        Subject     = $emailSubject
        Body        = $emailBody
        SmtpServer  = $smtpServer
        Attachments = $ArchiveFilePath
        BodyAsHtml  = $false
        ErrorAction = "Stop"
    }
    Send-MailMessage @mailParams

    Write-Host "All steps completed successfully." -ForegroundColor Green
    exit 0
}
catch {
    # Log full exception details
    $details = $_.Exception | Format-List * -Force | Out-String
    "Exception caught:`n$details" | Out-File -FilePath $LogFile -Append
    exit 1
}
finally {
    if ($session) {
        Remove-SFTPSession -SessionId $session.SessionId -ErrorAction SilentlyContinue
    }
    Stop-Transcript -ErrorAction SilentlyContinue
}
