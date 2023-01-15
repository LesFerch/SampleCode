# TimeRemaining

This is a package that provides an on-screen countdown of time remaining before a forced logout.

The execution is set up via two scheduled tasks, one running in the user context that provides the display, and one running in the SYSTEM context that will restore the countdown if the user kills the display task.

The display task includes an **Add Time** button that, when clicked, will run another script elevated to allow an administrator to give the user more time.

The script also display a toast notification to the user when there's 2 minutes left (the time can be adjusted in the script).

## Installation

Unzip the files to the folder **C:\HTA-TimeRemaining**. If you wish to use a different folder, you will have to edit the scheduled task XML files to the correct path.

**Note**: The folder cannot have any spaces in its name. This is a limitation of the **Elevate.exe** program needed to elevate and wait for the AddTime script.

Edit the permissions of the folder so that standard users have read-only access.

Open Task Scheduler, click **Import Task...** and select the file **TimeRemaining.xml**. Repeat for the file  **TimeRemaining-Maintain.xml**.

That's it! The scripts will run on next login.

**Note**: The tasks are set to run the scripts on login. Since **TimeRemaining.hta** is run in the user context, it automatically closes when the user logs out. However the **Maintain.vbs** runs as SYSTEM, so it detects when the user logs out and closes itself.

**Note**: The package is set up to run all the scripts in 32bit mode on a 64bit OS. There's no need to run in 64bit mode (that I know of) but if you wish to do so, change all paths, in all files, from **C:\Windows\SysWOW64** to **C:\Windows\System32** and replace **elevate.exe** with the 64bit version available [here](https://www.winability.com/elevate/).

If you wish to modify the package to run on a 32bit OS, make the same path change noted above, but keep the 32bit version of **elevate.exe**.
