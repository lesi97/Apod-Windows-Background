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