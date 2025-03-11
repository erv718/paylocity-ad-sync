# Paylocity to Active Directory Sync Engine

HR kept exporting employee data from Paylocity, and we had no automated way to get it into Active Directory. New hires sat in limbo, terminations stayed enabled, and rehires got missed entirely.

This repo contains two scripts that work together to solve it.

## How It Works

1. Paylocity automatically deposits a bi-weekly Excel export to an AWS Transfer Family SFTP endpoint.
2. A scheduled task runs `Paylocity_Report.ps1`, which connects via SSH key auth (Posh-SSH), downloads the file, renames it to `Paylocity_Data_Latest.xlsx`, archives a dated copy, and emails it to the ops team.
3. A second scheduled task runs `Sync-PaylocityToAD.ps1` with `-Apply`, which reads the downloaded file and handles the full employee lifecycle in AD.

Nobody touches a file.

## What Sync-PaylocityToAD.ps1 Does

- Reads `Paylocity_Data_Latest.xlsx`, splits by status (Active/Terminated)
- Creates new AD accounts with correct OU placement (Corporate vs Clubs)
- Disables and moves terminated users to a Disabled OU
- Detects rehires and re-enables their accounts
- Assigns Google Workspace license groups based on department
- Stores personal email in `extensionAttribute2` for offboarding
- Normalizes messy Paylocity data: duplicate headers, footer rows, Role IDs
- Outputs `SyncReport_*.xlsx` with tabs for each action taken
- Went through 13 versions over 5 months as edge cases kept surfacing

## Usage

```powershell
# Step 1 - Download latest export from SFTP
.\Paylocity_Report.ps1

# Step 2 - Dry run (no AD changes, just report)
.\Sync-PaylocityToAD.ps1 -Report "C:\Reports\Paylocity\Paylocity_Data_Latest.xlsx"

# Step 3 - Apply changes to AD
.\Sync-PaylocityToAD.ps1 -Report "C:\Reports\Paylocity\Paylocity_Data_Latest.xlsx" -Apply
```

## Requirements

- PowerShell 5.1+
- `ActiveDirectory` module (RSAT)
- `ImportExcel` module (auto-installs from PSGallery if missing)
- `Posh-SSH` module (`Install-Module Posh-SSH`)
- SSH private key configured for the SFTP endpoint
- AWS Transfer Family SFTP endpoint with Paylocity auto-deposit
- SMTP relay access for email notifications
- AD account with permissions to create/disable/move users

## Version History

The `archive/` folder contains all 13 iterations of `Sync-PaylocityToAD.ps1`. Each version added handling for a new edge case from real Paylocity exports.

## Blog Post

[13 Versions: Building a Paylocity to Active Directory Sync Engine](https://blog.soarsystems.cc/paylocity-hr2ad-sync/)
