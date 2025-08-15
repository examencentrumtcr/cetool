
README.TXT FOR BEHERENBESTANDEN.PS1

1. INTRODUCTION

This project was initiated to meet the need for a utility that could easily copy files to multiple computers.
It was named "Beherenbestanden" (Dutch for "Managing Files") because the script not only copies files, but also deletes and backs them up.

Later, the ability to move files between folders was added.
The project began in 2013, originally written in DOS language. Since then, the script has undergone many changes. The most significant update was the rewrite in PowerShell.

Main features:
- Copy files to folders
- Back up files

Additional features:
- Delete files
- Move or copy files between folders
- Explore destination folders
- Automatically update the script
- Log important tasks
- View log files
- Modify default personal settings
- Clean up old backup files
- Clean up log files at startup
- Show program info, README, and changelog
- Fetch a joke from moppenbot.nl

2. INSTALLATION

This script works on Windows 10 or later and requires PowerShell 5.1 or later.

To install the script, visit https://beherenbestanden.neveshuis.nl/setup and download "setup-beherenbestanden.exe".
Note: Microsoft Defender SmartScreen may block the file, but this warning can safely be ignored.

When the setup script starts, you’ll be presented with several installation options:
- Change the installation directory
- Create a desktop shortcut
- Perform a clean installation

It’s recommended to leave all options at their default values and simply click "Installeren" to proceed.

Alternatively, you can download the latest version as a ZIP file from:
http://beherenbestanden.neveshuis.nl/updates/latest
Extract it to a folder where the user has write permissions.

Note: This method requires you to manually create a shortcut (if desired) and be aware of Windows restrictions on PowerShell scripts.

3. CONFIGURATION

Configuration is only needed if:
- You did not create a desktop shortcut
- PowerShell is restricted from running scripts by default

To create a shortcut, simply run "snelkoppeling_maken.exe". This will also allow the script to run in unrestricted mode.

If you only want to enable PowerShell script execution, run "Wijzig_Executionpolicy_bypass.bat". This sets PowerShell to bypass mode.

If neither of these tools is used, or if an alternative method is applied, PowerShell may prompt you with an Execution Policy warning and ask whether to continue.

4. OPERATION

If you created a shortcut, just click it to run the script.

If not, right-click the script and choose "Open with" > "Windows PowerShell", or choose "Run with PowerShell".

5. FILE MANIFEST

This project consists of the following files and directories:

- Beherenbestanden.ps1 – Main startup script
- Beheren.ico – Program icon
- Snelkoppeling_maken.exe – Tool to create a desktop shortcut
- Wijzig_Executionpolicy_bypass.bat – Tool to set PowerShell Execution Policy to bypass
- Changelog.txt – List of changes per version
- Readme.txt – This file
- PNG – Directory containing icons and images used by the script

6. LICENSE NOTICE

This file is part of Beherenbestanden.

Beherenbestanden is free software: you can redistribute and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

Beherenbestanden is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.

See the GNU General Public License for more details.
You should have received a copy of the GNU GPL along with Beherenbestanden.
If not, see <https://www.gnu.org/licenses/>.

7. KNOWN BUGS

7.1. The script may crash while performing tasks.
However, the background process continues, and after a while, a message appears stating that the task has been completed.
If this happens, it's best to wait for the task to complete and verify that the script is still working.

Note: This issue was resolved starting with version 4.2.2. See the changelog at:
https://beherenbestanden.neveshuis.nl

8. AUTHORS AND ACKNOWLEDGMENTS

Main program author:
Benvindo Neves – https://neveshuis.nl/over-mij

Layout and images:
Marco Laluan

Thanks to:
Rene de Bruin, Edgar Seedorf, Robby Tosasi, and Rob Schortemeijer for their contributions to the project.

9. CHANGELOG

View the changelog at: https://beherenbestanden.neveshuis.nl
