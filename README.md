This is just a simple script to set your Windows wallpaper background to the <a href="https://apod.nasa.gov/apod/astropix.html">Nasa Astronomy Picture of the Day (APOD)</a>.

The script will create a scheduled task targeting itself to run on logon so place the script somewhere it won't be accidently deleted. On first run this will request administrator privileges to create the task. You can create the scheduled task manually to avoid this if preferred.

The scheduled task will be found in Task Scheduler Library > Apod and can be altered if required to point the script at a different path or to run on a different schedule.

By default this will save the APOD to C:\Users\%USERNAME%\Pictures\Apod.
This folder is never automatically emptied and can grow quite large in size if you wish to keep each picture for future backgrounds.

There can be instances where the APOD is instead a video, to handle this I've set a previous APOD I liked to be used instead.