# Requires: ImportExcel module, Git installed and configured

param(
    [string]$ExcelPath = "DevOps_Study_Tracker.xlsx"
)

if (-not (Test-Path $ExcelPath)) {
    Write-Error "Excel file not found at $ExcelPath"
    exit
}

# Prompt for task update
$day = Read-Host "Enter the Day number to update (e.g., 1)"
$status = Read-Host "Enter the new status (e.g., Done, In Progress, Skipped)"

# Load Excel and update the sheet
$data = Import-Excel -Path $ExcelPath -WorksheetName "Daily Learning Plan"

# Find the row matching the Day number
$row = $data | Where-Object { $_.Day -eq [int]$day }
if ($row) {
    $row.Status = $status
    $data | Export-Excel -Path $ExcelPath -WorksheetName "Daily Learning Plan" -Force
    Write-Host "Updated Day $day status to '$status' in Excel."
} else {
    Write-Error "Day $day not found in the tracker."
    exit
}

# Git commit and push
$commitMessage = "Update: Day $day marked as $status"
git add $ExcelPath
git commit -m $commitMessage
git push

Write-Host "Changes committed and pushed to GitHub."