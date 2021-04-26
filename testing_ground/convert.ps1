# Version 1

<#
Write-Output "The current directory is $(Get-Location)"
Write-Output "The curent date and time is $(Get-Date -UFormat "%T-%A-%d-%B-%Y")"
Write-Output "The curent date and time is $(Get-Date -UFormat "%H%M-%A-%d-%B-%Y")"
#>

Invoke-Expression -Command "C:\Users\drift\AppData\Local\Programs\Python\Python39\Scripts\pyuic5.exe -x 'app.ui' -o 'app_$(Get-Date -UFormat "%H%M-%A-%d-%B-%Y").py'"