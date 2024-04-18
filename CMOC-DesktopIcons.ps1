$programs = @{
    "Adobe Acrobat"  = "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
    "Excel"          = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
    "Google Chrome"  = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    "Microsoft Edge" = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    "Outlook"        = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
    "Word"           = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.exe"                  
}

$shortcuts = @(
    @{ Name = "Referral Spreadsheet"; URL = "https://lisalarkinmd-my.sharepoint.com/:x:/p/carrie/EaZiJHoidflIort-0aIshd8B72U9N_K_uvB-OjxquIJCsA?e=2Qm6VU" },
    @{ Name = "Rep Product List"; URL = "https://lisalarkinmd-my.sharepoint.com/:x:/p/karen/Ef-X4KhuoShFl06Fc4Dfg2MBqINrMpBt2aUmTdLnnxtyRA?e=fE7Ogw" }
)

#Check for shortcuts on user Desktop and in All Users desktop, if program is available and the shortcut isn't... Then recreate the shortcut on users desktop
#if not already present in ALl Users desktop folder
# Public Desktop path
$PublicDesktopPath = [Environment]::GetFolderPath("CommonDesktopDirectory")
$programs.GetEnumerator() | ForEach-Object {
    if (Test-Path -Path $_.Value) {
        if (-not (Test-Path -Path "$($PublicDesktopPath)\$($_.Key).lnk")) {
            Write-Host ("Shortcut for {0} not found in {1}, creating it now..." -f $_.Key, $_.Value)
            $shortcut = "$($PublicDesktopPath)\$($_.Key).lnk"
            $target = $_.Value
            $description = $_.Key
            $workingdirectory = (Get-ChildItem $target).DirectoryName
            $WshShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut($shortcut)
            $Shortcut.TargetPath = $target
            $Shortcut.Description = $description
            $shortcut.WorkingDirectory = $workingdirectory
            $Shortcut.Save()
        }
    }
}

# Loop through each shortcut
foreach ($shortcut in $shortcuts) {
    $name = $shortcut.Name
    $url = $shortcut.URL
    $iconPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    $iconIndex = "0"
    # Define shortcut path
    $shortcutPath = "$PublicDesktopPath\$name.url"
    # Check if shortcut already exists
    if (Test-Path -Path $shortcutPath) {
        Write-Host "Shortcut for $name already exists at $shortcutPath"
    }
    else {
        # Create Internet shortcut file content
        $shortcutContent = "[InternetShortcut]`r`nURL=$url`r`nIconFile=$iconPath`r`nIconIndex=$iconIndex"
        # Write content to shortcut file
        Set-Content -Path $shortcutPath -Value $shortcutContent
        Write-Host "Shortcut for $name created at $shortcutPath"
    }
}