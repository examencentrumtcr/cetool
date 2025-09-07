# Cetool

A tool to convert Eduarte exam files into Facet format

## 1. Introduction

Cetool is designed for vocational education schools in the Netherlands that use **Facet** to prepare digital examinations.

Facet requires specific data—such as student information, groups, exam types, and exam dates. This information is retrieved from the student administration system **Eduarte** as an Excel file. Since Facet can only read XML files, the Excel file must be converted. Cetool automates this process.

Additionally, Facet only accepts groups of students from a single school. To support examinations across multiple schools, Cetool splits the Eduarte Excel file into multiple XML files, with each file containing only students from one school.

### Key Features (v1 – released September 2025)

* **Excel → XML conversion**

  * Converts an Eduarte Excel file into multiple XML files, organized per exam.
  * Students with special accommodations (e.g., dyslexia) are placed in a separate XML file so they can be granted extra time.

* **Logging**

  * Records events in a logbook.
  * The logbook can be cleared manually or automatically remove entries older than a set number of days at startup.

* **Preferences**

  * Saves user settings such as:

    * Input directory for Excel files
    * Output directory for ZIP files
    * Output file name
    * Automatic logbook cleanup preferences
    * Console close behavior

* **Automatic updates**

  * Retrieves updates directly from GitHub.
  * Administrators can also download prerelease versions.

---

## 2. Installation

### System Requirements

* Windows 10 or later
* PowerShell 5.1 or later

### Installation via Setup (recommended)

1. Download the installer from:
   [Setup folder on GitHub](https://github.com/examencentrumtcr/cetool/tree/main/setup)
2. Extract `setup-cetool.zip`.
3. Run `setup-cetool.exe`.

> ⚠️ Note: Windows SmartScreen may warn you when running the installer. This happens because the file is not signed. If downloaded from the official repository, it can be safely run.

During installation, you can:

* Change the installation directory
* Create a desktop shortcut
* Perform a clean installation

We recommend keeping the default options and clicking **Installeren** to proceed.

### Manual Installation

1. Download the latest release from:
   [Release folder on GitHub](https://github.com/examencentrumtcr/cetool/tree/main/release)
2. Extract the ZIP file into a folder with write permissions.

> Note: With manual installation, you must create your own shortcut and manage PowerShell execution policies.

---

## 3. Configuration

Configuration is only required if:

* No desktop shortcut was created, or
* PowerShell script execution is restricted.

### Create a Desktop Shortcut

Run `snelkoppeling_maken.ps1` in PowerShell. This will also set the required execution policy.

### Adjust Execution Policy

Alternatively, run `Wijzig_Executionpolicy_bypass.bat` to set PowerShell to *Bypass* mode.

If neither tool is used, PowerShell may prompt with an execution policy warning when running Cetool.

---

## 4. Usage

* If you created a shortcut: **double-click it** to start Cetool.
* Without a shortcut: right-click the script and choose **Open with > Windows PowerShell** or **Run with PowerShell**.

---

## 5. File Manifest

The project contains the following files and directories:

```
Cetool.ps1                        Main startup script
Cetool_icoon.ico                  Application icon
Logboek.png                       Images used by the script
start.png                         Images used by the script
stoppen.png                       Images used by the script
Snelkoppeling_maken.ps1           Tool to create a desktop shortcut
Wijzig_Executionpolicy_bypass.bat Tool to set PowerShell Execution Policy to bypass
Changelog.txt                     List of version changes
Readme.txt                        This documentation
License                           GNU General Public License agreement
```

---

## 6. License

Cetool is licensed under the **GNU General Public License v3.0 (GPLv3)**.

This program is free software: you can redistribute and/or modify it under the terms of the GPL as published by the Free Software Foundation.
It is distributed **without any warranty**, not even for merchantability or fitness for a particular purpose.

See the [GNU GPL v3.0](https://www.gnu.org/licenses/gpl-3.0.html) for details.

---

## 7. Known Issues

Currently, there are no known issues.
If you encounter a bug, please report it via [GitHub Issues](https://github.com/examencentrumtcr/cetool/issues).

---

## 8. Authors and Acknowledgments

**Main Author**

* [Benvindo Neves](https://neveshuis.nl/over-mij)

**Contributors**

* Rob Schortemeijer

---

## 9. Changelog

See the [Changelog file](https://github.com/examencentrumtcr/cetool/tree/main/Changelog.txt) for details.

