:: Script om de Executionpolicy te wijzigen 

@echo off
Echo De ExecutionPolicy wordt op ByPass gezet zodat PowerShellscripts zonder problemen kunnen werken.

powershell -command "& {Set-ExecutionPolicy -ExecutionPolicy ByPass -scope currentuser -Force}"

Echo Uitgevoerd.
Echo Dit is de nieuwe waarde :

powershell -command "& {Get-ExecutionPolicy }

Echo.
pause