<# 
.SYNOPSIS
  Backup files from Desktop, Documents, and Downloads to local drives OR network shares.
  Creates subfolders by extension type and generates a CSV log of all copied files.
  PowerShell 7.x required (uses ForEach-Object -Parallel).

.DESCRIPTION
  - Scans Desktop, Documents, Downloads for specified file extensions
  - Supports BOTH local drives (C:, D:, E:) AND network shares (\\server\share)
  - Creates backup\[extension] folders automatically
  - COPIES files (doesn't move/delete originals)
  - Handles duplicate filenames by appending numbers
  - Generates CSV log with source and destination paths
  - Tests network connectivity before backup
  - Works across Windows languages via Environment.SpecialFolder
  - Has CPU throttle function, you can adjust it depending on the number of cores
  - Per default it will save all the files with extension you specified only from /desktop /download /documents folders

.NOTES
  Version: 0.4
  - Author : Michael DALLA RIVA, with the help of some AI
  - Release date : 26-Oct-2025
  - GitHub repository : https://github.com/michaeldallariva/PSBackupMyFiles
#>

#Requires -Version 7.0

#region ===== USER CONFIGURATION =====

# ┌────────────────────────────────────────────────────────────────┐
# │ BACKUP DESTINATION - CHOOSE ONE OF THE FOLLOWING:              │
# ├────────────────────────────────────────────────────────────────┤
# │ Option 1: LOCAL DRIVE                                          │
# │   Examples: "C:\Backup", "D:\MyBackups", "E:\Archives"         │
# │                                                                │
# │ Option 2: NETWORK SHARE (UNC path)                             │
# │   Examples: "\\192.168.1.10\mybackup"                          │
# │             "\\SERVER-NAME\BackupShare"                        │
# │             "\\192.168.1.100\SharedFolder\MyBackups"           │
# │                                                                │
# │ NOTE: For network shares, ensure:                              │
# │   • The share is accessible from your computer                 │
# │   • You have read/write permissions                            │
# │   • The path uses double backslashes (\\) at the start         │
# └────────────────────────────────────────────────────────────────┘

# 👇 SET YOUR BACKUP DESTINATION HERE 👇
# $BackupRoot = "\\192.168.1.10\backup$\"    # Alternative: Network share by IP
$BackupRoot = "C:\Backup"                    # Default: Local C: drive
# $BackupRoot = "D:\Backup"                  # Alternative: Local D: drive
# $BackupRoot = "E:\MyBackups"               # Alternative: Local E: drive
# $BackupRoot = "\\SERVER-NAME\BackupShare"  # Alternative: Network share by name

# ┌─────────────────────────────────────────────────────────────────┐
# │ FILE EXTENSIONS TO BACKUP                                       │
# └─────────────────────────────────────────────────────────────────┘
# Target extensions for backup (case-insensitive)
# Files like FILE.PDF, file.pdf, File.Pdf will all be matched
$Extensions = @(
  # Documents
  '.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.csv', '.txt', '.pub'
  # eBooks
  '.epub', '.mobi', '.azw', '.azw3'
  # Add more extensions as needed
)

# ┌────────────────────────────────────────────────────────────────┐
# │ SPECIFIC FOLDERS TO BACKUP (OPTIONAL)                          │
# └────────────────────────────────────────────────────────────────┘
# Specific folders to backup entirely (preserves full folder structure)
# Add full paths to folders you want to backup completely
# Uncomment and modify as needed:
$FoldersToBackup = @(
  # "C:\Users\YourName\Desktop\CV"
  # "C:\Users\YourName\Desktop\Scripts"
  # "C:\Users\YourName\Desktop\ThisFolder"
)

# ┌─────────────────────────────────────────────────────────────────┐
# │ PERFORMANCE SETTINGS                                            │
# └─────────────────────────────────────────────────────────────────┘
# Cap parallelism for file copying (auto-adjusts based on CPU cores)
$Throttle = [Math]::Min([Environment]::ProcessorCount, 4)

#endregion ===== END USER CONFIGURATION =====


#region ===== BACKUP DESTINATION VALIDATION =====
Write-Host "====================================" -ForegroundColor Cyan
Write-Host "  File Backup Script v0.4" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

# Create dated backup folder name
$BackupDate = (Get-Date).ToString('dd-MM-yyyy')
$BackupFolder = Join-Path $BackupRoot $BackupDate

# Determine if destination is network or local
$isNetworkPath = $BackupRoot -match '^\\\\[^\\]+'

if ($isNetworkPath) {
  Write-Host "🌐 Network backup destination detected" -ForegroundColor Cyan
  Write-Host "   Target: $BackupRoot" -ForegroundColor Gray
  Write-Host ""
  Write-Host "Testing network connectivity..." -ForegroundColor Yellow
  
  # Extract server/host from UNC path (\\server\share -> server)
  if ($BackupRoot -match '^\\\\([^\\]+)') {
    $networkHost = $Matches[1]
    
    # Test if host is reachable
    try {
      $pingResult = Test-Connection -ComputerName $networkHost -Count 1 -Quiet -ErrorAction Stop
      
      if ($pingResult) {
        Write-Host "✓ Network host '$networkHost' is reachable" -ForegroundColor Green
      } else {
        Write-Error "✗ Cannot reach network host '$networkHost'"
        Write-Host ""
        Write-Host "Troubleshooting tips:" -ForegroundColor Yellow
        Write-Host "  1. Verify the server/NAS is powered on" -ForegroundColor Gray
        Write-Host "  2. Check your network connection" -ForegroundColor Gray
        Write-Host "  3. Ensure the IP address or hostname is correct" -ForegroundColor Gray
        Write-Host "  4. Try accessing '$BackupRoot' in File Explorer" -ForegroundColor Gray
        return
      }
    } catch {
      Write-Warning "Could not test connectivity to '$networkHost'"
      Write-Host "Attempting to continue anyway..." -ForegroundColor Yellow
    }
    
    # Test if share is accessible
    Write-Host "Testing share access..." -ForegroundColor Yellow
    try {
      $testAccess = Test-Path -LiteralPath $BackupRoot -PathType Container
      if (-not $testAccess) {
        Write-Error "✗ Cannot access network share: $BackupRoot"
        Write-Host ""
        Write-Host "Troubleshooting tips:" -ForegroundColor Yellow
        Write-Host "  1. Verify you have read/write permissions to the share" -ForegroundColor Gray
        Write-Host "  2. Check if the share name is correct" -ForegroundColor Gray
        Write-Host "  3. Try mapping the network drive first" -ForegroundColor Gray
        Write-Host "  4. Ensure network credentials are configured" -ForegroundColor Gray
        Write-Host "  5. Try accessing '$BackupRoot' in File Explorer manually" -ForegroundColor Gray
        return
      }
      Write-Host "✓ Network share is accessible" -ForegroundColor Green
    } catch {
      Write-Error "✗ Failed to access network share: $BackupRoot"
      Write-Error $_.Exception.Message
      Write-Host ""
      Write-Host "Please verify the share path and your permissions." -ForegroundColor Yellow
      return
    }
  }
} else {
  Write-Host "💾 Local drive backup destination detected" -ForegroundColor Cyan
  Write-Host "   Target: $BackupRoot" -ForegroundColor Gray
}

Write-Host ""
#endregion


#region ===== CREATE BACKUP DIRECTORY =====
# Create main backup root if needed
if (-not (Test-Path -LiteralPath $BackupRoot -PathType Container)) {
  try {
    New-Item -Path $BackupRoot -ItemType Directory -Force | Out-Null
    Write-Host "✓ Created backup root directory: $BackupRoot" -ForegroundColor Green
  } catch {
    Write-Error "Failed to create backup root directory: $BackupRoot"
    Write-Error $_.Exception.Message
    if ($isNetworkPath) {
      Write-Host ""
      Write-Host "For network shares, ensure:" -ForegroundColor Yellow
      Write-Host "  • The parent folder exists on the network share" -ForegroundColor Gray
      Write-Host "  • You have write permissions" -ForegroundColor Gray
    }
    return
  }
}

# Create today's dated backup folder
if (-not (Test-Path -LiteralPath $BackupFolder -PathType Container)) {
  try {
    New-Item -Path $BackupFolder -ItemType Directory -Force | Out-Null
    Write-Host "✓ Created dated backup folder: $BackupDate" -ForegroundColor Green
  } catch {
    Write-Error "Failed to create dated backup folder: $BackupFolder"
    Write-Error $_.Exception.Message
    return
  }
} else {
  Write-Host "✓ Using existing backup folder: $BackupDate" -ForegroundColor Green
}
Write-Host ""
#endregion


#region ===== RESOLVE SOURCE FOLDERS =====
Write-Host "Resolving source folders..." -ForegroundColor Cyan

$desktop   = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
$documents = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)
$downloads = [Environment]::GetFolderPath('UserProfile') | Join-Path -ChildPath 'Downloads'

# Fallback for Downloads
if (-not (Test-Path -LiteralPath $downloads -PathType Container)) {
  try {
    $shell = New-Object -ComObject Shell.Application
    $downloads = $shell.NameSpace('shell:Downloads').Self.Path
  } catch {
    Write-Warning "Could not locate Downloads folder."
  }
}

$SourceRoots = @($desktop, $documents, $downloads) |
  Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Container) } |
  ForEach-Object { $_.TrimEnd('\') } |
  Sort-Object -Unique

if (-not $SourceRoots) {
  Write-Warning "No source folders (Desktop/Documents/Downloads) could be resolved. Exiting."
  return
}

Write-Host "Source folders:" -ForegroundColor Green
$SourceRoots | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }
Write-Host ""
#endregion


#region ===== CREATE EXTENSION SUBFOLDERS =====
Write-Host "Creating extension subfolders..." -ForegroundColor Cyan
$ExtensionsNorm = $Extensions | ForEach-Object { $_.ToLowerInvariant().TrimStart('.') }

foreach ($ext in $ExtensionsNorm) {
  $folderPath = Join-Path $BackupFolder $ext
  if (-not (Test-Path -LiteralPath $folderPath -PathType Container)) {
    try {
      New-Item -Path $folderPath -ItemType Directory -Force | Out-Null
      Write-Host "  ✓ Created: $BackupDate\$ext" -ForegroundColor Gray
    } catch {
      Write-Warning "Could not create folder for extension: $ext"
    }
  }
}
Write-Host ""
#endregion


#region ===== BACKUP SPECIFIC FOLDERS =====
if ($FoldersToBackup -and $FoldersToBackup.Count -gt 0) {
  Write-Host "Backing up specific folders..." -ForegroundColor Cyan
  $folderBackupLog = [System.Collections.Generic.List[pscustomobject]]::new()
  
  foreach ($sourceFolder in $FoldersToBackup) {
    if (-not (Test-Path -LiteralPath $sourceFolder -PathType Container)) {
      Write-Warning "Folder not found, skipping: $sourceFolder"
      continue
    }
    
    # Get folder name to create under backup
    $folderName = Split-Path -Leaf $sourceFolder
    $destFolder = Join-Path $BackupFolder $folderName
    
    try {
      Write-Host "  Processing: $folderName..." -ForegroundColor Gray
      
      # Create destination folder if it doesn't exist
      if (-not (Test-Path -LiteralPath $destFolder)) {
        New-Item -Path $destFolder -ItemType Directory -Force | Out-Null
      }
      
      # Get all files from source
      $sourceFiles = Get-ChildItem -LiteralPath $sourceFolder -Recurse -File
      $copiedCount = 0
      $skippedCount = 0
      
      foreach ($sourceFile in $sourceFiles) {
        # Calculate relative path
        $relativePath = $sourceFile.FullName.Substring($sourceFolder.Length).TrimStart('\')
        $destFilePath = Join-Path $destFolder $relativePath
        $destFileDir = Split-Path -Parent $destFilePath
        
        # Create subdirectory if needed
        if (-not (Test-Path -LiteralPath $destFileDir)) {
          New-Item -Path $destFileDir -ItemType Directory -Force | Out-Null
        }
        
        # Check if file exists and is identical
        $shouldCopy = $true
        if (Test-Path -LiteralPath $destFilePath) {
          $existingFile = Get-Item -LiteralPath $destFilePath
          if ($existingFile.Length -eq $sourceFile.Length -and 
              $existingFile.LastWriteTime -eq $sourceFile.LastWriteTime) {
            # File is identical, skip
            $shouldCopy = $false
            $skippedCount++
          }
        }
        
        if ($shouldCopy) {
          Copy-Item -LiteralPath $sourceFile.FullName -Destination $destFilePath -Force -ErrorAction Stop
          # Preserve modification date
          (Get-Item -LiteralPath $destFilePath).LastWriteTime = $sourceFile.LastWriteTime
          $copiedCount++
        }
      }
      
      $folderBackupLog.Add([pscustomobject]@{
        SourceFolder = $sourceFolder
        DestinationFolder = $destFolder
        FilesCopied = $copiedCount
        FilesSkipped = $skippedCount
        BackupDate = Get-Date
        Status = 'Success'
      })
      
      Write-Host "    ✓ Copied: $copiedCount files, Skipped: $skippedCount files" -ForegroundColor Green
      
    } catch {
      Write-Warning "Failed to backup folder: $sourceFolder - $($_.Exception.Message)"
      $folderBackupLog.Add([pscustomobject]@{
        SourceFolder = $sourceFolder
        DestinationFolder = 'FAILED'
        FilesCopied = 0
        FilesSkipped = 0
        BackupDate = Get-Date
        Status = "Failed: $($_.Exception.Message)"
      })
    }
  }
  
  Write-Host ""
}
#endregion


#region ===== SCAN FOR FILES =====
Write-Host "Scanning for files to backup..." -ForegroundColor Cyan
$sw = [System.Diagnostics.Stopwatch]::StartNew()

$Work = [System.Collections.Generic.List[pscustomobject]]::new()
foreach ($root in $SourceRoots) {
  foreach ($ext in $Extensions) {
    $Work.Add([pscustomobject]@{ Root = $root; Ext = $ext })
  }
}

$foundFiles = $Work | ForEach-Object -Parallel {
  $root = $_.Root
  $ext  = $_.Ext

  try {
    Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue |
      Where-Object { $_.Extension -ieq $ext } |
      Select-Object Name, FullName, Extension, Length, LastWriteTime, DirectoryName
  } catch {
    # Silent error handling
  }
} -ThrottleLimit $Throttle

$FilesToBackup = @($foundFiles) | Where-Object { $_ }
$sw.Stop()

Write-Host "✓ Found $($FilesToBackup.Count) files to backup in $([math]::Round($sw.Elapsed.TotalSeconds, 2))s" -ForegroundColor Green
Write-Host ""

if ($FilesToBackup.Count -eq 0) {
  Write-Warning "No files found to backup. Exiting."
  return
}
#endregion


#region ===== COPY FILES WITH DUPLICATE HANDLING =====
Write-Host "Starting backup process..." -ForegroundColor Cyan
$sw.Restart()

# Thread-safe collection for logging
$syncHash = [hashtable]::Synchronized(@{
  CopiedFiles = [System.Collections.Concurrent.ConcurrentBag[pscustomobject]]::new()
  SuccessCount = 0
  FailCount = 0
  SkipCount = 0
})

$FilesToBackup | ForEach-Object -Parallel {
  $file = $_
  $backupFolder = $using:BackupFolder
  $sync = $using:syncHash
  
  try {
    # Determine destination folder based on extension
    $ext = $file.Extension.ToLowerInvariant().TrimStart('.')
    $destFolder = Join-Path $backupFolder $ext
    $destPath = Join-Path $destFolder $file.Name
    
    # Check if file already exists with same size and modification date
    if (Test-Path -LiteralPath $destPath) {
      $existingFile = Get-Item -LiteralPath $destPath
      
      # Compare size and last write time
      if ($existingFile.Length -eq $file.Length -and 
          $existingFile.LastWriteTime -eq $file.LastWriteTime) {
        # File is identical, skip it
        $sync.CopiedFiles.Add([pscustomobject]@{
          SourcePath = $file.FullName
          DestinationPath = $destPath
          FileName = $file.Name
          Extension = $file.Extension
          SizeKB = [math]::Round($file.Length / 1KB, 2)
          OriginalDate = $file.LastWriteTime
          BackupDate = Get-Date
          Status = 'Skipped (already exists, unchanged)'
        })
        
        [System.Threading.Interlocked]::Increment([ref]$sync.SkipCount) | Out-Null
        return
      }
      
      # File exists but is different, create new version with number suffix
      $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
      $extension = $file.Extension
      $counter = 1
      
      do {
        $newName = "${baseName}_$counter$extension"
        $destPath = Join-Path $destFolder $newName
        $counter++
      } while (Test-Path -LiteralPath $destPath)
    }
    
    # Copy the file
    Copy-Item -LiteralPath $file.FullName -Destination $destPath -Force -ErrorAction Stop
    
    # Preserve original modification date
    (Get-Item -LiteralPath $destPath).LastWriteTime = $file.LastWriteTime
    
    # Log successful copy
    $sync.CopiedFiles.Add([pscustomobject]@{
      SourcePath = $file.FullName
      DestinationPath = $destPath
      FileName = $file.Name
      Extension = $file.Extension
      SizeKB = [math]::Round($file.Length / 1KB, 2)
      OriginalDate = $file.LastWriteTime
      BackupDate = Get-Date
      Status = 'Success'
    })
    
    [System.Threading.Interlocked]::Increment([ref]$sync.SuccessCount) | Out-Null
    
  } catch {
    # Log failed copy
    $sync.CopiedFiles.Add([pscustomobject]@{
      SourcePath = $file.FullName
      DestinationPath = 'FAILED'
      FileName = $file.Name
      Extension = $file.Extension
      SizeKB = [math]::Round($file.Length / 1KB, 2)
      OriginalDate = $file.LastWriteTime
      BackupDate = Get-Date
      Status = "Failed: $($_.Exception.Message)"
    })
    
    [System.Threading.Interlocked]::Increment([ref]$sync.FailCount) | Out-Null
  }
} -ThrottleLimit $Throttle

$sw.Stop()
$CopiedFiles = @($syncHash.CopiedFiles)
#endregion


#region ===== GENERATE CSV LOG =====
$timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
$csvPathFiles = Join-Path $BackupFolder "backup_log_files_$timestamp.csv"

$CopiedFiles | 
  Sort-Object Status, Extension, FileName |
  Export-Csv -LiteralPath $csvPathFiles -NoTypeInformation -Encoding UTF8

# Save folder backup log if any folders were backed up
if ($folderBackupLog -and $folderBackupLog.Count -gt 0) {
  $csvPathFolders = Join-Path $BackupFolder "backup_log_folders_$timestamp.csv"
  $folderBackupLog | 
    Export-Csv -LiteralPath $csvPathFolders -NoTypeInformation -Encoding UTF8
}

Write-Host ""
Write-Host "====================================" -ForegroundColor Cyan
Write-Host "  Backup Complete!" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Statistics:" -ForegroundColor Green
Write-Host "  Files by extension:" -ForegroundColor Cyan
Write-Host "    ✓ Successfully copied: $($syncHash.SuccessCount) files" -ForegroundColor Green
Write-Host "    ⊘ Skipped (unchanged): $($syncHash.SkipCount) files" -ForegroundColor Yellow
if ($syncHash.FailCount -gt 0) {
  Write-Host "    ✗ Failed: $($syncHash.FailCount) files" -ForegroundColor Red
}

if ($folderBackupLog -and $folderBackupLog.Count -gt 0) {
  $successfulFolders = ($folderBackupLog | Where-Object { $_.Status -eq 'Success' }).Count
  $totalFolderFilesCopied = ($folderBackupLog | Where-Object { $_.Status -eq 'Success' } | Measure-Object -Property FilesCopied -Sum).Sum
  $totalFolderFilesSkipped = ($folderBackupLog | Where-Object { $_.Status -eq 'Success' } | Measure-Object -Property FilesSkipped -Sum).Sum
  Write-Host "  Complete folders:" -ForegroundColor Cyan
  Write-Host "    ✓ Successfully copied: $totalFolderFilesCopied files from $successfulFolders folders" -ForegroundColor Green
  Write-Host "    ⊘ Skipped (unchanged): $totalFolderFilesSkipped files" -ForegroundColor Yellow
}

Write-Host "  ⏱ Duration: $([math]::Round($sw.Elapsed.TotalSeconds, 2)) seconds" -ForegroundColor Cyan
Write-Host ""
Write-Host "Backup location: $BackupFolder" -ForegroundColor Yellow
Write-Host "CSV log (files): $csvPathFiles" -ForegroundColor Yellow
if ($folderBackupLog -and $folderBackupLog.Count -gt 0) {
  Write-Host "CSV log (folders): $csvPathFolders" -ForegroundColor Yellow
}
Write-Host ""

# Show breakdown by extension
$breakdown = $CopiedFiles | 
  Where-Object { $_.Status -eq 'Success' } |
  Group-Object Extension | 
  Sort-Object Name

if ($breakdown) {
  Write-Host "Files backed up by extension:" -ForegroundColor Cyan
  foreach ($group in $breakdown) {
    $totalSize = ($group.Group | Measure-Object -Property SizeKB -Sum).Sum
    Write-Host "  $($group.Name): $($group.Count) files ($([math]::Round($totalSize/1024, 2)) MB)" -ForegroundColor Gray
  }
}

Write-Host ""
Write-Host "✓ Backup completed successfully!" -ForegroundColor Green
#endregion

# Open backup folder (works for both local and network paths)
try {
  Start-Process -FilePath explorer.exe -ArgumentList $BackupFolder
} catch {
  Write-Host "Note: Could not open backup folder automatically." -ForegroundColor Yellow
  Write-Host "You can access it manually at: $BackupFolder" -ForegroundColor Gray
}
