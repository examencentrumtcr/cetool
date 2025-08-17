<# script om een snelkoppeling te maken van Beherenbestanden met de juiste opties
#>

# Declareren variabelen ----------------------------------------------------------

# startmap van het programma
$startmap = $PSScriptRoot
# $startmap="$env:USERPROFILE\cetool"

# hoofd programma die uitgevoerd wordt
[string] $hoofdprognaam = "cetool.ps1"

# samenvoegen installatiemap met hoofdprogramma. nodig voor aanmaken snelkoppeling
$hoofdprog = -join ("$startmap","\","$hoofdprognaam")

# Locatie van de executeable van Powershell in een windows 64bit systeem. 
# Let op dat bij een 32bit systeem dit anders is.
[string] $locatiePSexe = "c:\windows\system32\WindowsPowerShell\v1.0\powershell.exe"

# Locatie van de snelkoppeling. Dit is de bureaublad map van de gebruiker. 
$linkbureaublad = [Environment]::GetFolderPath("Desktop")

# Naam en locatie van de snelkoppeling die gemaakt wordt
$ShortcutPath = "$linkbureaublad\CE-tool.lnk"

# Einde declareren variabelen ----------------------------------------------------

# Maken van snelkoppeling
$WScriptObj = New-Object -ComObject ("WScript.Shell")
$shortcut = $WscriptObj.CreateShortcut($ShortcutPath)
$shortcut.TargetPath = $locatiePSexe
# Gekozen om de snelkoppeling te maken zonder de optie -WindowStyle Hidden, zodat de gebruiker de console kan zien.
$shortcut.Arguments = "-ExecutionPolicy Bypass -NoProfile -file " + '"' + "$hoofdprog" + '"' 
# $shortcut.Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -NoProfile -file " + '"' + "$hoofdprog" + '"' 
$shortcut.IconLocation = "$startmap" + "\" + "cetool_icoon.ico"
$shortcut.WorkingDirectory = "$startmap"
$shortcut.Save()

# Einde maken van snelkoppeling

# Info geven
write-host ""
Write-Host "Een snelkoppeling is gemaakt naar het programma $hoofdprognaam met de volgende opties: "
Write-Host "Link naar programma   : $hoofdprog" 
Write-Host "Startmap              : $startmap"
write-host "Snelkoppeling gemaakt : $ShortcutPath "

# Wacht op een toets om af te sluiten.
write-host ""
write-host "Klaar. Druk op een toets om af te sluiten...."
$null = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
