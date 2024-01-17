
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    # Relaunch the script with administrator rights
    Start-Process powershell.exe -Verb RunAs -ArgumentList $MyInvocation.MyCommand.Definition
    Exit
}

Add-Type -AssemblyName System.Windows.Forms

$userDownloads = "${env:USERPROFILE}\Downloads"
cd $userDownloads
$userDocuments = "${env:USERPROFILE}\Documents"
$userPictures = "${env:USERPROFILE}\Pictures"
$pythonDirs = "${env:LOCALAPPDATA}\Programs\Python"

$pythonExe = Get-ChildItem -Path $pythonDirs -Filter "python.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1


if ($null -eq $pythonExe) {
    Write-Output "Python not installed, downloading Python from https://www.python.org/ftp/python/3.12.1/python-3.12.1-amd64.exe"
    $pythonDownloadLink = "https://www.python.org/ftp/python/3.12.1/python-3.12.1-amd64.exe"
    $pythonLocalName = Join-Path $userDownloads "Python-3.12.1.exe"
    Invoke-WebRequest -Uri $pythonDownloadLink -OutFile $pythonLocalName
    Start-Process -FilePath $pythonLocalName -Args "/quiet InstallAllUsers=0 InstallLauncherAllUsers=0 PrependPath=1 Include_test=0" -Wait -NoNewWindow
} 

$apodPicturesPath = Join-Path $userPictures "Apod"
if (-not (Test-Path -Path $apodPicturesPath)) {
    New-Item -Path $apodPicturesPath -ItemType Directory | Out-Null
}

$apodDocumentsPath = Join-Path $userDocuments "Apod"
if (-not (Test-Path -Path $apodDocumentsPath)) {
    New-Item -Path $apodDocumentsPath -ItemType Directory | Out-Null
}

cd $apodDocumentsPath

$apodPythonFile = @'
import http.client
import ctypes
import json
import urllib.request
from urllib.error import URLError, HTTPError
import os
from pathlib import Path
import re
import datetime
import ssl

currentDate = datetime.date.today()
apodPath = Path.home() / "Pictures/Apod"
apiKey = os.getenv('Apod')

def noApiError():
    WS_EX_TOPMOST = 0x40000
    windowTitle = "No API Key"
    message = "It doesn't appear as though there is an API key to use the NASA Api. You will need to add this into your operating system to work. "
    message += "Open the Run dialog and search for SystemPropertiesAdvanced "
    message += "Then click on Environment Variables > New "
    message += "Set the Variable Name as Apod and the value as your API key "
    ctypes.windll.user32.MessageBoxExW(None, message, windowTitle, WS_EX_TOPMOST)

if apiKey is None:
    noApiError()

# check if todays image has already been downloaded, if yes then quit out
def checkFileDates(apodPath):
    for fileName in os.listdir(apodPath):
        filePath = os.path.join(apodPath, fileName)
        if os.path.isfile(filePath):
            createdTimestamp = os.path.getctime(filePath)
            createdDate = datetime.date.fromtimestamp(createdTimestamp)
            createdTime = datetime.datetime.fromtimestamp(createdTimestamp).time()
            eightAm = datetime.time(8, 0)  # 8:00 AM

            if createdDate == currentDate and createdTime >= eightAm:
                return True
    return False

if checkFileDates(apodPath):
    exit()

def removeInvalidPathChars(img):
    invalidChars = '[<>:"/\\\\|?*]'
    newImgName = re.sub(invalidChars, '', img)
    return newImgName

print("Updating Wallpaper With Today's Astronomy Picture Which Is Of")

# Make API Request
apiKey = os.getenv('Apod')
conn = http.client.HTTPSConnection(
    "api.nasa.gov",
    context = ssl._create_unverified_context()
)
headersListJson = {
 "Accept": "application/json" 
}
payload = ""
conn.request("GET", "/planetary/apod?api_key=" + apiKey, payload, headersListJson)
response = conn.getresponse()
result = response.read()
data = json.loads(result.decode("utf-8"))

#ssl._create_default_https_context = ssl._create_unverified_context #Uncomment this line if you get an SSL error.
# Check if Apod today is video or image
if (data["media_type"] != "video"):
    imgName = data["title"]
    imgName = removeInvalidPathChars(imgName)
    imgURL = data["hdurl"]
    altImgURL = data["url"]
    print(imgName)
    # Download Apod image
    localPath = os.path.join(apodPath, imgName + ".jpg")
    try:
        urllib.request.urlretrieve(imgURL, localPath)
    except HTTPError as error:
        if error.code == 404:
            try:
                urllib.request.urlretrieve(altImgURL, localPath)
            except URLError as error:
                print(f"Failed to download the SD image. Error: {error}")
                imgName = "NGC 6960 The Witch's Broom Nebula"
                imgURL = "https://apod.nasa.gov/apod/image/1804/ngc6960_Pugh_2000.jpg"
                localPath = os.path.join(apodPath, imgName + ".jpg")
                urllib.request.urlretrieve(imgURL, localPath)
        else:
            print(f"HTTP Error: {error}")
    except URLError as error:
        print(f"Failed to download the HD image. Error: {error}")
else: 
    imgName = "NGC 6960 The Witch's Broom Nebula"
    imgURL = "https://apod.nasa.gov/apod/image/1804/ngc6960_Pugh_2000.jpg"

# Set image as background
class Main:
    def __init__(self):
        path = os.path.join(apodPath, imgName + ".jpg")
        ctypes.windll.user32.SystemParametersInfoW(20, 0, path , 0)
application = Main()

exit()
'@

$apodBatFile = @'
@echo off
py apod.py
'@

$pythonFilePath = Join-Path -Path $userDocuments\Apod -ChildPath "apod.py"
$batFilePath = Join-Path -Path $userDocuments\Apod -ChildPath "apod.bat"

New-Item -Path $pythonFilePath -Type File -Value $apodPythonFile -Force
New-Item -Path $batFilePath -Type File -Value $apodBatFile -Force

$userName = $Env:UserName
$pcName = $env:COMPUTERNAME

$taskTrigger = New-ScheduledTaskTrigger -AtLogOn
$taskAction = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c C:\Users\$username\Documents\Apod\apod.bat"
$taskPrincipal = New-ScheduledTaskPrincipal -UserId "$pcName\$userName" -LogonType Interactive -RunLevel Highest
Register-ScheduledTask "Apod" -Action $taskAction -Trigger $taskTrigger -Principal $taskPrincipal

$promptToInputApiKey = "You will now need to input your Nasa API key into your system variables`n`n"
$promptToInputApiKey += "Open the Run dialog by holding the Windows Key and press R, then search for SystemPropertiesAdvanced`n`n"
$promptToInputApiKey += "Click on Environment Variables > New`n`n"
$promptToInputApiKey += "Set the Variable Name as Apod and the value as your API key from https://api.nasa.gov/"

[System.Windows.Forms.MessageBox]::Show($promptToInputApiKey, "Finished")
