<#
 .SYNOPSIS
 Reinicia la cola de impresión

 .DESCRIPTION
 Script que para el servicio de cola de impresión, borra los archivos pendientes de imprimir de C:\Windows\system32\spool\PRINTERS y vuelve a arrancar el servicio

 .PARAMETER machine
 Nombre de la maquina en la reiniciar el servicio

 .EXAMPLE
 Clearspooler.ps1 ZOCHMGU0000
 #>

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True,Position=0)]
    [string]$machine
    )
    
    if (Test-NetConnection $machine -InformationLevel Quiet)
    {
    (Get-Service -ComputerName $machine -Name Spooler).Stop()
    Get-ChildItem \\$machine\C$\Windows\System32\spool\PRINTERS | foreach ($_) {remove-item $_.fullname}
    (Get-Service -ComputerName $machine -Name Spooler).Start()
    }
