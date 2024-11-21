$bookmarkFile = "$env:USERPROFILE\hop_bookmarks.txt"

function Initialize-Hop {
    param(
        [switch]$Force
    )

    # Check for administrative privileges
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $admin = [Security.Principal.WindowsPrincipal]::new($currentUser)
    
    if (-not $admin.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
        Write-Host "Error: Hop initialization requires administrator privileges." -ForegroundColor Red
        Write-Host "Please run PowerShell as an Administrator and try again." -ForegroundColor Yellow
        return
    }

    # Get the current script's directory
    $scriptDir = Split-Path -Parent $PSCommandPath

    # Check if the script is already in PATH
    $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    if ($currentPath -split ';' -contains $scriptDir -and -not $Force) {
        Write-Host "Hop script directory is already in the system PATH." -ForegroundColor Green
        return
    }

    try {
        # Modify the machine-level PATH
        $newPath = $currentPath + ";$scriptDir"
        [Environment]::SetEnvironmentVariable("Path", $newPath, "Machine")

        Write-Host "Hop script added to system PATH. Please restart your terminal or log off/log on for changes to take effect." -ForegroundColor Green
    }
    catch {
        Write-Host "Error adding Hop script to PATH: $_" -ForegroundColor Red
    }
}

function Load-Bookmarks {
    $bookmarks = @{}
    if (Test-Path $bookmarkFile) {
        $lines = Get-Content $bookmarkFile
        foreach ($line in $lines) {
            $parts = $line -split '\|'
            if ($parts.Length -eq 5) {
                $bookmarks[$parts[0]] = @{
                    Path         = $parts[1]
                    Category     = $parts[2]
                    LastAccessed = $parts[3]
                    AccessCount  = [int]$parts[4]
                }
            }
        }
    }
    else {
        Write-Host "No bookmark file found." -ForegroundColor Yellow
    }
    return $bookmarks
}

function Save-Bookmarks {
    param ($bookmarks)
    Remove-Item -Path $bookmarkFile -Force -ErrorAction SilentlyContinue
    $bookmarks.GetEnumerator() | ForEach-Object {
        "$($_.Key)|$($_.Value.Path)|$($_.Value.Category)|$($_.Value.LastAccessed)|$($_.Value.AccessCount)" | Out-File -FilePath $bookmarkFile -Append
    }
}

function Add-Bookmark {
    param ($name, $path, $category)

    if (!$category) {
        $category = "general"
    }

    Write-Debug "Adding bookmark: $name, $path, $category"
    $bookmarks = Load-Bookmarks
    if ($bookmarks[$name]) {
        Write-Host "Bookmark '$name' already exists. Use 'hop remove $name' to remove it." -ForegroundColor Red
        return
    }

    if (Test-Path $path) {
        $bookmarks[$name] = @{
            Path         = $path
            Category     = $category
            LastAccessed = ""
            AccessCount  = 0
        }
        Save-Bookmarks $bookmarks
        Write-Host "Bookmark '$name' added under category '$category' at $path" -ForegroundColor Green
    }
    else {
        Write-Host "Invalid path: $path" -ForegroundColor Red
    }
}

function Set-As-Bookmark {
    param ($name, $category)

    if (!$category) {
        $category = "general"
    }

    $bookmarks = Load-Bookmarks
    if ($bookmarks.ContainsKey($name)) {
        Write-Host "Bookmark '$name' already exists. Use 'hop remove $name' to remove it." -ForegroundColor Red
        return
    }

    $bookmarks[$name] = @{
        Path         = (Get-Location).Path
        Category     = $category
        LastAccessed = ""
        AccessCount  = 0
    }
    Save-Bookmarks $bookmarks
    Write-Host "Current location set as bookmark '$name' under category '$category' at $((Get-Location).Path)" -ForegroundColor Green
}

function Rename-Bookmark {
    param (
        $oldName,
        $newName,
        [switch]$Category,
        $newCategory
    )
    
    $bookmarks = Load-Bookmarks
    if (!$oldName) {
        Write-Host "Error: Specify the old bookmark name.'$oldName'" -ForegroundColor Red
        return
    }
    if ((!$newName) -and (!$newCategory)) {
        Write-Host "Error: Specify either a new name or a new category." -ForegroundColor Red
        return
    }
    
    if (-not $bookmarks.ContainsKey($oldName)) {
        Write-Host "Bookmark '$oldName' not found." -ForegroundColor Red
        return
    }
    
    if ($newName) {
        if ($bookmarks.ContainsKey($newName)) {
            Write-Host "Bookmark '$newName' already exists. Choose a different name." -ForegroundColor Red
            return
        }
        
        $bookmarks[$newName] = $bookmarks[$oldName]
        $bookmarks.Remove($oldName)
    }
    
    if ($Category) {
        $targetName = if ($newName) { $newName } else { $oldName }
        $bookmarks[$targetName].Category = $newCategory
    }
    
    Save-Bookmarks $bookmarks
    
    if ($newName -and $Category) {
        Write-Host "Bookmark '$oldName' renamed to '$newName' with new category '$newCategory'." -ForegroundColor Green
    }
    elseif ($newName) {
        Write-Host "Bookmark '$oldName' renamed to '$newName'." -ForegroundColor Green
    }
    elseif ($Category) {
        Write-Host "Category for bookmark '$oldName' changed to '$newCategory'." -ForegroundColor Green
    }
}

function Go-To-Bookmark {
    param ($name)
    $bookmarks = Load-Bookmarks
    if ($bookmarks.ContainsKey($name)) {
        Set-Location $bookmarks[$name].Path
        $bookmarks[$name].LastAccessed = Get-Date
        $bookmarks[$name].AccessCount++
        Save-Bookmarks $bookmarks
        Write-Host "Navigated to bookmark '$name' at $($bookmarks[$name].Path)" -ForegroundColor Cyan
    }
    else {
        Write-Host "Bookmark '$name' not found." -ForegroundColor Red
    }
}

function List-Bookmarks {
    param ($word, $category = "")
    $bookmarks = Load-Bookmarks
    $filtered = if ($category) {
        $bookmarks.GetEnumerator() | Where-Object { $_.Value.Category -eq $category }
    }
    elseif ($word) {
        $bookmarks.GetEnumerator() | Where-Object { $_.Key -like "*$word*" }
    }
    else {
        $bookmarks.GetEnumerator()
    }
    $filtered | ForEach-Object {
        Write-Host -NoNewline $_.Key -ForegroundColor Green
        Write-Host -NoNewline " "
        Write-Host $_.Value.Path -ForegroundColor Cyan
    }
}

function Show-Stats {
    $bookmarks = Load-Bookmarks
    if ($bookmarks.Count -eq 0) {
        Write-Host "No bookmarks found." -ForegroundColor Yellow
        return
    }

    $bookmarks.GetEnumerator() | ForEach-Object {
        $lastAccessed = if ($_.Value.LastAccessed) {
            [DateTime]$_.Value.LastAccessed -as [DateTime]
        } else {
            "Never"
        }

        [PSCustomObject]@{
            Bookmark      = $_.Key 
            Path          = $_.Value.Path
            Category      = $_.Value.Category
            'Last Access' = $lastAccessed
            'Access Count' = $_.Value.AccessCount
        }
    } | Format-Table -AutoSize -Wrap
}



function Show-Recent {
    $bookmarks = Load-Bookmarks
    $recent = $bookmarks.GetEnumerator() | Where-Object { $_.Value.LastAccessed -ne "" } | Sort-Object { $_.Value.LastAccessed } -Descending | Select-Object -First 10
    Write-Host "Recently Accessed Bookmarks:" -ForegroundColor Cyan
    $recent | ForEach-Object { Write-Host "$($_.Key) -> Last Accessed: $($_.Value.LastAccessed)" }
}

function Show-Frequent {
    $bookmarks = Load-Bookmarks
    $frequent = $bookmarks.GetEnumerator() | Sort-Object { $_.Value.AccessCount } -Descending | Select-Object -First 10
    Write-Host "Frequently Accessed Bookmarks:" -ForegroundColor Cyan
    $frequent | ForEach-Object { Write-Host "$($_.Key) -> Access Count: $($_.Value.AccessCount)" }
}

function Clear-Bookmarks {
    Write-Host "WARNING: This action will delete all bookmarks permanently!" -ForegroundColor Yellow
    $confirmation = Read-Host "Are you sure you want to proceed? Type 'yes' to confirm"
    
    if ($confirmation -eq "yes") {
        Remove-Item -Path $bookmarkFile -Force -ErrorAction SilentlyContinue
        Write-Host "All bookmarks cleared." -ForegroundColor Red
    } else {
        Write-Host "Operation canceled. No bookmarks were cleared." -ForegroundColor Green
    }
}


function Remove-Bookmark {
    param ($name)
    $bookmarks = Load-Bookmarks
    if ($bookmarks.ContainsKey($name)) {
        $bookmarks.Remove($name)
        Save-Bookmarks $bookmarks
        Write-Host "Bookmark '$name' removed." -ForegroundColor Yellow
    }
    else {
        Write-Host "Bookmark '$name' not found." -ForegroundColor Red
    }
}

function Show-Help {
    Write-Host "Hop Bookmark Management Help" -ForegroundColor Green
    Write-Host "`nBookmark Management:" -ForegroundColor Cyan
    Write-Host "  hop add <name> <path> [<category>]" -ForegroundColor White -NoNewline
    Write-Host "    - Add a new bookmark (default category: general)" 
    
    Write-Host "  hop set <name> [<category>]" -ForegroundColor White -NoNewline
    Write-Host "        - Set current location as a bookmark"
    
    Write-Host "  hop to <name>" -ForegroundColor White -NoNewline
    Write-Host "                   - Navigate to a saved bookmark"
    
    Write-Host "  hop rename <oldname> [<newname>] [-c <category>]" -ForegroundColor White -NoNewline
    Write-Host " - Rename bookmark or change its category"

    Write-Host "`nListing and Searching:" -ForegroundColor Cyan
    Write-Host "  hop list [<word>]" -ForegroundColor White -NoNewline
    Write-Host "               - List all bookmarks or search by keyword"
    
    Write-Host "  hop list -c <category>" -ForegroundColor White -NoNewline
    Write-Host "        - List bookmarks in a specific category"

    Write-Host "`nAnalytics and Management:" -ForegroundColor Cyan
    Write-Host "  hop stats" -ForegroundColor White -NoNewline
    Write-Host "                  - Display detailed bookmark statistics"
    
    Write-Host "  hop recent" -ForegroundColor White -NoNewline
    Write-Host "               - Show 10 most recently accessed bookmarks"
    
    Write-Host "  hop frequent" -ForegroundColor White -NoNewline
    Write-Host "           - Show 10 most frequently accessed bookmarks"
    
    Write-Host "  hop remove <name>" -ForegroundColor White -NoNewline
    Write-Host "         - Remove a specific bookmark"
    
    Write-Host "  hop clear" -ForegroundColor White -NoNewline
    Write-Host "                 - Clear all bookmarks (with confirmation)"

    Write-Host "`nUtility:" -ForegroundColor Cyan
    Write-Host "  hop help" -ForegroundColor White -NoNewline
    Write-Host "                  - Display this help information"

    Write-Host "`nTips:" -ForegroundColor Cyan
    Write-Host "  - Use quotes around paths or names with spaces" -ForegroundColor DarkGray
    Write-Host "  - Optional arguments are shown in square brackets" -ForegroundColor DarkGray
}
# Enhanced switch case with more robust error handling
switch ($args[0]) {
    { $_ -eq $null } { Show-Help }
    "help" { Show-Help }
    "add" { 
        if ($args.Length -ge 3) {
            try {
                Add-Bookmark -name $args[1] -path $args[2] -category $args[3] 
            }
            catch {
                Write-Host "Error adding bookmark: $_" -ForegroundColor Red
            }
        }
        else {
            Write-Host "Usage: hop add <name> <path> [<category>]" -ForegroundColor Yellow
        }
    }
    "set" { 
        if ($args.Length -ge 2) {
            try {
                Set-As-Bookmark -name $args[1] -category $args[2]
            }
            catch {
                Write-Host "Error setting bookmark: $_" -ForegroundColor Red
            }
        }
        else {
            Write-Host "Usage: hop set <name> [<category>]" -ForegroundColor Yellow
        }
    }
    "to" { 
        if ($args.Length -ge 2) {
            try {
                Go-To-Bookmark -name $args[1]
            }
            catch {
                Write-Host "Error navigating to bookmark: $_" -ForegroundColor Red
            }
        }
        else {
            Write-Host "Usage: hop to <name>" -ForegroundColor Yellow
        }
    }
    "list" {
        try {
            if ($args[1] -eq "-c" -and $args[2]) {
                List-Bookmarks -category $args[2]
            }
            else {
                List-Bookmarks $args[1]
            }
        }
        catch {
            Write-Host "Error listing bookmarks: $_" -ForegroundColor Red
        }
    }
    "stats" { Show-Stats }
    "recent" { Show-Recent }
    "init" { Initialize-Hop }
    "frequent" { Show-Frequent }
    "remove" { 
        if ($args.Length -ge 2) {
            try {
                Remove-Bookmark -name $args[1]
            }
            catch {
                Write-Host "Error removing bookmark: $_" -ForegroundColor Red
            }
        }
        else {
            Write-Host "Usage: hop remove <name>" -ForegroundColor Yellow
        }
    }
    "clear" { Clear-Bookmarks }
    "rename" { 
        $oldName = $args[1]

        if ($args.Contains("-c")) {
            # If "-c" is present, get newName (if available) and newCategory
            $newName = if ($args[2] -ne "-c") { $args[2] } elseif ($args[4]) { $args[4] } else { $null }
            $newCategory = $args[$args.IndexOf("-c") + 1]
        } 
        else {
            # If "-c" is not present, treat the second argument as newName
            $newName = $args[2]
            $newCategory = $null
        }

        # Validate oldName and call Rename-Bookmark with relevant arguments
        if ($oldName) {
            try {
                Rename-Bookmark -oldName $oldName -newName $newName -Category:($newCategory -ne $null) -newCategory $newCategory
            }
            catch {
                Write-Host "Error renaming bookmark: $_" -ForegroundColor Red
            }
        }
        else {
            Write-Host "Please specify at least an old name for the bookmark." -ForegroundColor Red
        }
    }
    default { 
        Write-Host "Unknown command. Use 'hop help' for instructions." -ForegroundColor Red 
    }
}

    
