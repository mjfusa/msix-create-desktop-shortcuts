# RunOnce Package Support Framework PowerShell script
# to create Desktop shortcuts
# Prerequisites: 
#     1) App must be packaged with an AppExecution alias. (Example here: https://gist.github.com/mjfusa/b7ad353ddb3c8f3b8127684d47b863fb)
#     AppExecution alias docs here: https://docs.microsoft.com/en-us/uwp/schemas/appxpackage/uapmanifestschema/element-uap5-appexecutionalias
#     2) App must be packaged with the Package Support framework and configured to run this script. (Example here: https://github.com/mjfusa/PSFCurrentWorkingDirectory)
# This script will:
# 1) Open manifest interate throught Applications collection extracting DisplayName, Executeable and Description elements for Shortcut parameters.
# 2) Extract icon from EXE and save to $env:APPDATA
# 3) Create shortcut, setting working directory %localappdata%\Microsoft\WindowsApps and save to $env:USERPROFILE\desktop\<DisplayName>.lnk

# Log will be written to C:\TEMP\CreateShortcuts.log
$res=Test-Path "C:\TEMP"
if ($res -eq $false)
{
    New-Item -Path "C:\TEMP" -ItemType direct
}
$LogFile = "C:\TEMP\CreateShortcuts.log"
Clear-Content $LogFile -ErrorAction SilentlyContinue
Function Log-Write
{
   Param ($Text)
 
   Add-Content $LogFile -Value $Text
}

# Get Package Name - should be in root of package
$ScriptPath = $pwd.Path #Split-Path -Path $MyInvocation.MyCommand.Path

# Application parameters used to create shortcut
$AppObject = [PSCustomObject] @{
    DisplayName = ''
    Executable  = ''
    Description = ''
}
$appArray = [System.Collections.ArrayList]@()

# Get DisplayName, Description from manifest
[xml]$manifest = Get-Content ($ScriptPath + "\AppxManifest.xml")
$appsInManifest = $manifest.Package.Applications.Application
foreach ($app in $appsInManifest) {
    # Don't create a shortcut if the app is hidden from the Start menu
    if ($app.VisualElements.AppListEntry -eq "none") {
        continue
    }
    $object = [PSCustomObject] @{
        DisplayName = ''
        Executable  = ''
        Description = ''
    }
    $object.DisplayName = $app.VisualElements.DisplayName;
    $object.Description = $app.VisualElements.Description;
    $appArray.Add($object);
}

# Get Executable from config.json
$config = Get-Content ($ScriptPath + "\config.json") | ConvertFrom-Json
$cnt = 0
foreach ($app in $config.applications) {
    $appArray[$cnt].Executable = $app.executable
    $cnt++;
}
Function CreateShortcut {
    param (
        $app = [PSCustomObject] @{
            DisplayName = ''
            Executable  = ''
            Description = ''
        },
        [string]$destinationDesktopPath
    )
    $WshShell = New-Object -comObject WScript.Shell
    $lnkPath = ($destinationDesktopPath + "\" + $app.DisplayName + ".lnk")
    $Shortcut = $WshShell.CreateShortcut($lnkPath)
    $Shortcut.TargetPath = $(Split-Path $app.Executable -Leaf)
    $Shortcut.IconLocation = ($env:APPDATA + "\" + $app.DisplayName + ".ico")
    $Shortcut.WorkingDirectory = ($ENV:LOCALAPPDATA + "\Microsoft\WindowsApps")
    $Shortcut.Description = ($app.Description)
    $Shortcut.Save()
}

$desinationIconPath = $env:APPDATA
# Setup correct desktop folder for shortcut destination
$destinationDesktopPath = $env:USERPROFILE + "\Desktop"
if ($null -ne $env:OneDriveCommercial) {
    if (Test-Path($env:OneDriveCommercial + "\Desktop")) {
        $destinationDesktopPath = $env:OneDriveCommercial + "\Desktop"        
    }
}

Function CopyIcon
{
    Param ( 
        [Parameter(Mandatory = $true)]
        [string]$iconSourcePath,
        [string]$iconDestinationPath
    )

    Copy-Item -Path $iconSourcePath -Destination $iconDestinationPath
}

$date= Get-Date
Log-Write -Text $("Create shorcuts: " + $date)
Log-Write -Text $("Destination desktop path for shortcut: " + $destinationDesktopPath)
Foreach ($app in $appArray) {
    try {
        CopyIcon $("Icons\" + $app.DisplayName + ".ico") $desinationIconPath 
    } catch {
        Log-Write $Error[0]
    }
    if (Test-Path($($desinationIconPath + "\" + $app.DisplayName + ".ico")) ) {
        Log-Write -Text $("Icon successfully copied: " + $desinationIconPath + "\" + $app.DisplayName + ".ico")
    } else {
        Log-Write -Text $("Icon NOT copied: " + $desinationIconPath + "\" + $app.DisplayName + ".ico")
    }
    CreateShortcut -app $app -destinationDesktopPath $destinationDesktopPath
    if (Test-Path($($destinationDesktopPath + "\" + $app.DisplayName + ".lnk"))) {
        Log-Write -Text $("Shortcut (lnk) successfully created: " + $destinationDesktopPath + "\" + $app.DisplayName + ".lnk")
    } else {
        Log-Write -Text $("Shortcut (lnk) NOT created: " + $destinationDesktopPath + "\" + $app.DisplayName + ".lnk")
    }
}
