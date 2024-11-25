$taskName = "Apod"
$scheduledTaskExists = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

if ($scheduledTaskExists -eq $null) {

    $scriptPath = $MyInvocation.MyCommand.Path

    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
        Start-Sleep 60
        exit
    }

    $action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c start /min powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -Hidden

    Register-ScheduledTask -Action $action -Trigger $trigger -TaskName $taskName
} 

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class Wallpaper {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@
$SPI_SETDESKWALLPAPER = 0x0014
$SPIF_UPDATEINIFILE = 0x01
$SPIF_SENDCHANGE = 0x02

$RegPath = "HKCU:\Control Panel\Desktop"
$apodSaveFolder = "$env:USERPROFILE\Pictures\Apod\"

if (-not (Test-Path -Path $apodSaveFolder)) {
    New-Item -Path "$env:USERPROFILE\Pictures\" -Name "Apod" -ItemType "directory"
}

$currentDate = Get-Date -Format "dd-MM-yyyy"
$savedImagesToday = Get-ChildItem -Path $apodSaveFolder -Filter "*$currentDate.jpg"

if ($savedImagesToday.Count -eq 0) { 
    $headers = @{
        "Accept" = "application/json"
    }
    $reqUrl = "https://api.nasa.gov/planetary/apod?api_key=DEMO_KEY"
    $response = Invoke-RestMethod -Uri $reqUrl -Method Get -Headers $headers  

    if ($response.media_type -eq "image" -and $response.hdurl -ne $null -and $response.title -ne $null) {
        $imageName = $response.title
        $imageName = $response.title -replace "[\\/:*?`"<>|]+", ""
        $sdImage = $response.url
        $hdImage = $response.hdurl
        $backgroundimage = Join-Path -Path $apodSaveFolder -ChildPath ($imageName + " - " + $currentDate + ".jpg")

        try {
			Invoke-WebRequest $hdImage -OutFile $backgroundimage
		} catch {
			Write-Output "HD image download failed, likely the URL provided doesnt work."
            Write-Output "Attempting to download the SD image instead."
			Invoke-WebRequest $sdImage -OutFile $backgroundimage
		}
        [Wallpaper]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $backgroundimage, $SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE)

    }
    else {
        Write-Host "No image today. Maybe a nice video instead."

        $backupImageName = "The Cone Nebula from Hubble"
        $backupImageURL = "https://apod.nasa.gov/apod/image/2402/cone_hubbleschmidt_4048.jpg"
        $backupBackgroundimage = Join-Path -Path $apodSaveFolder -ChildPath ($backupImageName + "_" + $currentDate + ".jpg")

        Invoke-WebRequest $backupImageURL -OutFile $backupBackgroundimage
        [Wallpaper]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $backupBackgroundimage, $SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE)
    }
}
else {
    $savedImage = Join-Path -Path $apodSaveFolder -ChildPath ($savedImagesToday.Name)
    [Wallpaper]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $savedImage, $SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE)
    Write-Host "Image for today already exists."
}