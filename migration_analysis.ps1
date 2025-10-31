# ============================================
# Data Migration Pre-Analysis Script
# 2025/10/31
# ============================================

$sourcePath = "D:\"
$targetPath = "Z:\"
$reportPath = "D:\migration_report"

# Create report folder
New-Item -ItemType Directory -Path $reportPath -Force | Out-Null

Write-Host "=== Starting Data Collection ===" -ForegroundColor Cyan
Write-Host "Source: $sourcePath"
Write-Host "Target: $targetPath"
Write-Host ""

# Define $files array to store results from Step 1 for reuse in later steps
$files = @() 

# 1. Basic Statistics & File Collection
Write-Host "[1/8] Calculating total file count and size (Including Subfolders)..." -ForegroundColor Yellow
$fileStats = @{
    TotalFiles = 0
    TotalSizeGB = 0
    AvgFileSizeKB = 0
}

try {
    # --- Start of Progress Bar Implementation ---
    $fileCounter = 0
    $totalLength = [long]0
    $startTime = Get-Date
    
    # Get all files and calculate stats, using ForEach-Object to show progress
    $files = Get-ChildItem -Path "$sourcePath\*.csv" -File -Recurse | ForEach-Object {
        $file = $_ # Capture the current file object
        
        $fileCounter++
        $totalLength += $file.Length
        
        # Update progress bar and fixed-position console status for every 100 files
        if ($fileCounter % 100 -eq 0 -or $fileCounter -lt 10) {
            $elapsedTime = (Get-Date) - $startTime
            $rate = $fileCounter / $elapsedTime.TotalSeconds
            
            # 1. Update standard Write-Progress (top of console)
            Write-Progress -Activity "File Statistics Analysis (Step 1/8)" `
                           -Status "Scanning disk... Processed files: $($fileCounter.ToString('N0'))" `
                           -CurrentOperation "Rate: $($rate.ToString('N0')) files/sec"

            # 2. Fixed-position console progress (using carriage return \r and -NoNewLine)
            $message = "Current Count: $($fileCounter.ToString('N0')) | Rate: $($rate.ToString('N0')) files/sec.           "
            Write-Host "`r$message" -NoNewLine
        }
        
        # Pass the file object to be collected in $files for reuse
        $_
    }
    
    # Ensure cursor moves to a new line after the fixed-position updates stop
    Write-Host ""
    
    # Complete the progress bar after iteration finishes
    Write-Progress -Activity "File Statistics Analysis (Step 1/8)" -Status "Analysis Complete" -Completed

    if ($fileCounter -gt 0) {
        $fileStats.TotalFiles = $fileCounter
        $fileStats.TotalSizeGB = [math]::Round($totalLength / 1GB, 2)
        $fileStats.AvgFileSizeKB = [math]::Round($totalLength / $fileCounter / 1KB, 2)
    }
    
    Write-Host "  Total Files: $($fileStats.TotalFiles)"
    Write-Host "  Total Size: $($fileStats.TotalSizeGB) GB"
    Write-Host "  Average File Size: $($fileStats.AvgFileSizeKB) KB"
} catch {
    Write-Host "  錯誤: $_" -ForegroundColor Red
}

$totalFiles = $files.Count # Total file count for percentage calculation in subsequent steps

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
# Reuse $files from Step 1 for efficiency
$sample = $files |
    Select-Object -First 1000 |
    Select-Object Name, Length, CreationTime, LastWriteTime, Directory

$sample | Export-Csv "$reportPath\file_sample.csv" -NoTypeInformation -Encoding UTF8

# 4. Date Distribution
Write-Host "[4/8] Analyzing date distribution (Including Subfolders)..." -ForegroundColor Yellow
$i = 0
$dateGroupsData = $files | ForEach-Object {
    $i++
    $percentage = [math]::Round(($i / $totalFiles) * 100, 0)
    
    # 1. Update standard Write-Progress (top of console)
    Write-Progress -Activity "Date Distribution Analysis (Step 4/8)" `
                   -Status "Processing... $percentage% Complete ($i/$totalFiles items)" `
                   -PercentComplete $percentage
    
    # 2. Fixed-position console progress (using carriage return \r and -NoNewLine) - update every 100 items
    if ($i % 100 -eq 0 -or $i -eq $totalFiles) {
        $message = "Progress: $percentage% ($i/$totalFiles items) processed.           "
        Write-Host "`r$message" -NoNewLine
    }

    [PSCustomObject]@{ DateKey = $_.LastWriteTime.ToString("yyyy-MM") }
}

# Ensure cursor moves to a new line after the fixed-position updates stop
Write-Host ""
Write-Progress -Activity "Date Distribution Analysis (Step 4/8)" -Status "Analysis Complete" -Completed

$dateGroups = $dateGroupsData |
    Group-Object DateKey |
    Select-Object @{Name='YearMonth';Expression={$_.Name}}, Count |
    Sort-Object YearMonth

$dateGroups | Export-Csv "$reportPath\date_distribution.csv" -NoTypeInformation -Encoding UTF8
Write-Host "  Time Range: $($dateGroups[0].YearMonth) 到 $($dateGroups[-1].YearMonth)"

# 5. Size Distribution
Write-Host "[5/8] Analyzing file size distribution (Including Subfolders)..." -ForegroundColor Yellow
$i = 0
$sizeGroupsData = $files | ForEach-Object {
    $i++
    $percentage = [math]::Round(($i / $totalFiles) * 100, 0)
    
    # 1. Update standard Write-Progress (top of console)
    Write-Progress -Activity "Size Distribution Analysis (Step 5/8)" `
                   -Status "Processing... $percentage% Complete ($i/$totalFiles items)" `
                   -PercentComplete $percentage

    # 2. Fixed-position console progress (using carriage return \r and -NoNewLine) - update every 100 items
    if ($i % 100 -eq 0 -or $i -eq $totalFiles) {
        $message = "Progress: $percentage% ($i/$totalFiles items) processed.           "
        Write-Host "`r$message" -NoNewLine
    }

    $sizeBucket = switch ($_.Length) {
        {$_ -lt 10KB} { "< 10KB" }
        {$_ -lt 100KB} { "10-100KB" }
        {$_ -lt 1MB} { "100KB-1MB" }
        {$_ -lt 10MB} { "1-10MB" }
        default { "> 10MB" }
    }
    [PSCustomObject]@{ SizeBucket = $sizeBucket }
}

# Ensure cursor moves to a new line after the fixed-position updates stop
Write-Host ""
Write-Progress -Activity "Size Distribution Analysis (Step 5/8)" -Status "Analysis Complete" -Completed

$sizeGroups = $sizeGroupsData | Group-Object SizeBucket | Select-Object Name, Count

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
$i = 0
$fileNameList = $files | ForEach-Object {
    $i++
    $percentage = [math]::Round(($i / $totalFiles) * 100, 0)
    
    # 1. Update standard Write-Progress (top of console)
    Write-Progress -Activity "Duplicate Name Check (Step 7/8)" `
                   -Status "Processing... $percentage% Complete ($i/$totalFiles items)" `
                   -PercentComplete $percentage

    # 2. Fixed-position console progress (using carriage return \r and -NoNewLine) - update every 100 items
    if ($i % 100 -eq 0 -or $i -eq $totalFiles) {
        $message = "Progress: $percentage% ($i/$totalFiles items) processed.           "
        Write-Host "`r$message" -NoNewLine
    }

    [PSCustomObject]@{ FileName = $_.Name }
}

# Ensure cursor moves to a new line after the fixed-position updates stop
Write-Host ""
Write-Progress -Activity "Duplicate Name Check (Step 7/8)" -Status "Analysis Complete" -Completed

$duplicates = $fileNameList |
    Group-Object FileName |
    Where-Object { $_.Count -gt 1 }

if ($duplicates) {
    Write-Host "  Found $($duplicates.Count) sets of duplicate filenames" -ForegroundColor Red
    $duplicates | Export-Csv "$reportPath\duplicate_names.csv" -NoTypeInformation -Encoding UTF8
} else {
    Write-Host "  No duplicate filenames found" -ForegroundColor Green
}

# 8. Generate Summary Report
Write-Host "[8/8] Generating summary report..." -ForegroundColor Yellow
# Ensure $dateGroups is not null before accessing its elements
$earliestDate = if ($dateGroups.Count -gt 0) { $dateGroups[0].YearMonth } else { "N/A" }
$latestDate = if ($dateGroups.Count -gt 0) { $dateGroups[-1].YearMonth } else { "N/A" }
$duplicateCount = if ($duplicates) { $duplicates.Count } else { 0 }

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
- Earliest: $earliestDate
- Latest: $latestDate

Duplicate Names: $duplicateCount sets

Detailed reports saved to: $reportPath
========================================
"@

$summary | Out-File "$reportPath\summary.txt" -Encoding UTF8
Write-Host $summary

Write-Host "`n=== Data Collection Complete ===" -ForegroundColor Green
Write-Host "Report Location: $reportPath"
