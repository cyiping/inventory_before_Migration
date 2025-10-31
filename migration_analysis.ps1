# ============================================
# Data Migration Pre-Analysis Script
# ============================================

$sourcePath = "D:\"
$reportPath = "D:\migration_report"

# Create report folder
New-Item -ItemType Directory -Path $reportPath -Force | Out-Null

Write-Host "=== Starting Data Collection ===" -ForegroundColor Cyan
Write-Host "Source: $sourcePath"
Write-Host "Target: $targetPath"
Write-Host ""

# 1. Basic Statistics
Write-Host "[1/8] Calculating total file count and size (Including Subfolders)..." -ForegroundColor Yellow
$fileStats = @{
    TotalFiles = 0
    TotalSizeGB = 0
    AvgFileSizeKB = 0
}

try {
    # Added -Recurse to include all subdirectories
    $measure = Get-ChildItem -Path "$sourcePath\*.csv" -File -Recurse | 
        Measure-Object -Property Length -Sum -Average
    
    $fileStats.TotalFiles = $measure.Count
    $fileStats.TotalSizeGB = [math]::Round($measure.Sum / 1GB, 2)
    $fileStats.AvgFileSizeKB = [math]::Round($measure.Average / 1KB, 2)
    
    Write-Host "  Total Files: $($fileStats.TotalFiles)"
    Write-Host "  Total Size: $($fileStats.TotalSizeGB) GB"
    Write-Host "  Average File Size: $($fileStats.AvgFileSizeKB) KB"
} catch {
    Write-Host "  Error: $_" -ForegroundColor Red
}

# 2. Folder Structure
Write-Host "[2/8] Analyzing folder structure..." -ForegroundColor Yellow
$folderInfo = Get-ChildItem -Path $sourcePath -Directory -Recurse -ErrorAction SilentlyContinue |
    Select-Object FullName, @{
        Name='RelativePath'
        Expression={$_.FullName.Replace($sourcePath, '')}
    }

Write-Host "  Number of Subfolders: $($folderInfo.Count)"
$folderInfo | Export-Csv "$reportPath\folder_structure.csv" -NoTypeInformation -Encoding UTF8

# 3. Sampling Analysis
Write-Host "[3/8] Sampling file analysis (Top 1000, including subfolders)..." -ForegroundColor Yellow
# Added -Recurse to include all subdirectories
$sample = Get-ChildItem -Path "$sourcePath\*.csv" -File -Recurse |
    Select-Object -First 1000 |
    Select-Object Name, Length, CreationTime, LastWriteTime, Directory

$sample | Export-Csv "$reportPath\file_sample.csv" -NoTypeInformation -Encoding UTF8

# 4. Date Distribution
Write-Host "[4/8] Analyzing date distribution (Including Subfolders)..." -ForegroundColor Yellow
# Added -Recurse to include all subdirectories
$dateGroups = Get-ChildItem -Path "$sourcePath\*.csv" -File -Recurse |
    Group-Object {$_.LastWriteTime.ToString("yyyy-MM")} |
    Select-Object @{Name='YearMonth';Expression={$_.Name}}, Count |
    Sort-Object YearMonth

$dateGroups | Export-Csv "$reportPath\date_distribution.csv" -NoTypeInformation -Encoding UTF8
Write-Host "  Time Range: $($dateGroups[0].YearMonth) to $($dateGroups[-1].YearMonth)"

# 5. Size Distribution
Write-Host "[5/8] Analyzing file size distribution (Including Subfolders)..." -ForegroundColor Yellow
# Added -Recurse to include all subdirectories
$sizeGroups = Get-ChildItem -Path "$sourcePath\*.csv" -File -Recurse |
    Group-Object {
        switch ($_.Length) {
            {$_ -lt 10KB} { "< 10KB" }
            {$_ -lt 100KB} { "10-100KB" }
            {$_ -lt 1MB} { "100KB-1MB" }
            {$_ -lt 10MB} { "1-10MB" }
            default { "> 10MB" }
        }
    } | Select-Object Name, Count

$sizeGroups | Format-Table -AutoSize
$sizeGroups | Export-Csv "$reportPath\size_distribution.csv" -NoTypeInformation -Encoding UTF8

# 6. Disk Space Check
Write-Host "[6/8] Checking disk space..." -ForegroundColor Yellow
$sourceDrive = Get-PSDrive D
$targetDrive = Get-PSDrive Z

$spaceCheck = [PSCustomObject]@{
    SourceUsed_GB = [math]::Round($sourceDrive.Used / 1GB, 2)
    TargetFree_GB = [math]::Round($targetDrive.Free / 1GB, 2)
    SpaceSufficient = ($targetDrive.Free -gt ($sourceDrive.Used * 1.1))  # Reserve 10% extra space
}

$spaceCheck | Format-List
$spaceCheck | Export-Csv "$reportPath\space_check.csv" -NoTypeInformation -Encoding UTF8

if (-not $spaceCheck.SpaceSufficient) {
    Write-Host "  Warning: Target disk space is insufficient!" -ForegroundColor Red
}

# 7. Duplicate Filename Check
Write-Host "[7/8] Checking for duplicate filenames (Including Subfolders)..." -ForegroundColor Yellow
# Added -Recurse to include all subdirectories
$duplicates = Get-ChildItem -Path "$sourcePath\*.csv" -File -Recurse |
    Group-Object Name |
    Where-Object { $_.Count -gt 1 }

if ($duplicates) {
    Write-Host "  Found $($duplicates.Count) sets of duplicate filenames" -ForegroundColor Red
    $duplicates | Export-Csv "$reportPath\duplicate_names.csv" -NoTypeInformation -Encoding UTF8
} else {
    Write-Host "  No duplicate filenames found" -ForegroundColor Green
}

# 8. Generate Summary Report
Write-Host "[8/8] Generating summary report..." -ForegroundColor Yellow
$summary = @"
========================================
Data Migration Pre-Analysis Report
========================================
Analysis Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Source Path: $sourcePath
Target Path: $targetPath

Basic Statistics:
- Total Files (Including Subfolders): $($fileStats.TotalFiles)
- Total Size (Including Subfolders): $($fileStats.TotalSizeGB) GB
- Average File Size (Including Subfolders): $($fileStats.AvgFileSizeKB) KB
- Number of Subfolders: $($folderInfo.Count)

Disk Space:
- Source Used: $($spaceCheck.SourceUsed_GB) GB
- Target Available: $($spaceCheck.TargetFree_GB) GB
- Space Sufficient: $($spaceCheck.SpaceSufficient)

Time Range:
- Earliest: $($dateGroups[0].YearMonth)
- Latest: $($dateGroups[-1].YearMonth)

Duplicate Names: $($duplicates.Count) sets

Detailed reports saved to: $reportPath
========================================
"@

$summary | Out-File "$reportPath\summary.txt" -Encoding UTF8
Write-Host $summary

Write-Host "`n=== Data Collection Complete ===" -ForegroundColor Green
Write-Host "Report Location: $reportPath"
