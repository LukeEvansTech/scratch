# Comprehensive Ivanti Registry & Filesystem Search
# Run as Administrator

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "IVANTI REGISTRY & FILESYSTEM SEARCH" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# ========================================
# REGISTRY SEARCH - All Ivanti entries
# ========================================
Write-Host "`n### REGISTRY SEARCH ###`n" -ForegroundColor Yellow

Write-Host "=== HKLM:\SOFTWARE (checking all subkeys) ===" -ForegroundColor Green
Get-ChildItem -Path "HKLM:\SOFTWARE" -ErrorAction SilentlyContinue | 
Where-Object { $_.Name -like "*Ivanti*" } | 
ForEach-Object {
    Write-Host "`n[REG] $($_.Name)" -ForegroundColor Cyan
    Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
}

Write-Host "`n=== HKLM:\SOFTWARE\Wow6432Node (32-bit) ===" -ForegroundColor Green
Get-ChildItem -Path "HKLM:\SOFTWARE\Wow6432Node" -ErrorAction SilentlyContinue | 
Where-Object { $_.Name -like "*Ivanti*" } | 
ForEach-Object {
    Write-Host "`n[REG] $($_.Name)" -ForegroundColor Cyan
    Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
}

Write-Host "`n=== HKCU:\SOFTWARE (Current User) ===" -ForegroundColor Green
Get-ChildItem -Path "HKCU:\SOFTWARE" -ErrorAction SilentlyContinue | 
Where-Object { $_.Name -like "*Ivanti*" } | 
ForEach-Object {
    Write-Host "`n[REG] $($_.Name)" -ForegroundColor Cyan
    Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
}

# ========================================
# FILESYSTEM SEARCH - All Ivanti directories
# ========================================
Write-Host "`n`n### FILESYSTEM SEARCH ###`n" -ForegroundColor Yellow

$search_paths = @(
    "$env:ProgramFiles\Ivanti",
    "${env:ProgramFiles(x86)}\Ivanti",
    "$env:ProgramData\Ivanti",
    "$env:LocalAppData\Ivanti",
    "C:\Ivanti",
    "C:\Program Files\Ivanti*",
    "C:\Program Files (x86)\Ivanti*"
)

foreach ($path in $search_paths) {
    if (Test-Path $path) {
        Write-Host "`n=== Checking: $path ===" -ForegroundColor Green
        
        # List all files recursively
        $files = Get-ChildItem -Path $path -Recurse -ErrorAction SilentlyContinue
        
        Write-Host "Total items found: $($files.Count)" -ForegroundColor Cyan
        
        # Show directory structure
        Write-Host "`n[DIRECTORIES]" -ForegroundColor Magenta
        $files | Where-Object { $_.PSIsContainer } | ForEach-Object {
            Write-Host "  📁 $($_.FullName)"
        }
        
        # Show all files
        Write-Host "`n[FILES]" -ForegroundColor Magenta
        $files | Where-Object { !$_.PSIsContainer } | Sort-Object Extension | ForEach-Object {
            Write-Host "  📄 $($_.Name) - $($_.FullName) [$($_.Length) bytes]"
        }
        
        # Show config-type files
        Write-Host "`n[CONFIG/XML FILES]" -ForegroundColor Magenta
        $files | Where-Object { $_.Extension -match '\.(xml|ini|conf|config|json)$' } | ForEach-Object {
            Write-Host "  ⚙️  $($_.FullName)" -ForegroundColor Yellow
            Write-Host "     Size: $($_.Length) bytes | Modified: $($_.LastWriteTime)"
        }
    }
    else {
        Write-Host "`n[NOT FOUND] $path" -ForegroundColor DarkGray
    }
}

# ========================================
# Check common Ivanti config locations
# ========================================
Write-Host "`n`n### CHECKING COMMON IVANTI CONFIG FILES ###`n" -ForegroundColor Yellow

$config_files = @(
    "$env:ProgramData\Ivanti\config.xml",
    "$env:ProgramData\Ivanti\settings.ini",
    "$env:ProgramFiles\Ivanti\config.xml",
    "${env:ProgramFiles(x86)}\Ivanti\config.xml",
    "$env:LocalAppData\Ivanti\config.xml"
)

foreach ($file in $config_files) {
    if (Test-Path $file) {
        Write-Host "`n[FOUND CONFIG] $file" -ForegroundColor Green
        Write-Host "Content (first 50 lines):" -ForegroundColor Cyan
        Get-Content $file -Head 50 | ForEach-Object { Write-Host "  $_" }
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "SEARCH COMPLETE" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan