 <#
  Script que genera un objeto con información del PC.
  Modelo, nº de serie, Version del sistema operativo, disco y espacio libre, RAM, procesador y grafica
  .PARAMETER machine
  nombre del equipo o ip
  #>

[CmdletBinding()]
Param(
  [Parameter(Mandatory=$True,Position=1)]
   [string]$machine
   )

  $ADComment = (Get-ADComputer $machine -Properties Description).Description
  if (Test-NetConnection $machine -InformationLevel Quiet)
    {
      $direccionip = (Resolve-DnsName $machine).IPAddress
      $modelo = (Get-CimInstance -ComputerName $machine -ClassName Win32_ComputerSystem).Model
      if ($modelo -match '^All\sSeries|^\s*$')
        {
        $modelo = (Get-WmiObject -ComputerName $machine Win32_BaseBoard).Manufacturer + " " + (Get-WmiObject -ComputerName $machine Win32_BaseBoard).Product
        }
      $nserie =  (Get-CimInstance -ComputerName $machine -ClassName Win32_SystemEnclosure).SerialNumber
      if ($nserie.Equals("Chassis Serial Number"))
        {
        $nserie = (Get-WmiObject -ComputerName $machine Win32_BaseBoard).SerialNumber
        }
      $Inventario = (cat \\$machine\c$\Inventario\Inventario.txt | Select-string num_oc).Line.split("|")[1]
      $MAC = (Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName $machine).MACAddress
      $SPEED =  (Get-WmiObject -Class Win32_networkAdapter).speed
        if ($SPEED -ge 1000000000) {$interface = "Gigabit"}
        elseif ($SPEED -lt 1000000000 -ge 100000000) {$interface = "FastEth"}
        elseif ($SPEED -lt 100000000 -ge 10000000) {$interface = "Ethernet"}
        else {$interface = "OTRA"}
      $SO = (Get-WmiObject -ComputerName $machine -ClassName Win32_OperatingSystem).Version
      $Logindate = [Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject -ComputerName $machine Win32_OperatingSystem).LastBootUpTime)
      $Disco = (Get-WmiObject win32_logicalDisk -ComputerName $machine -Filter "DeviceID='C:'").Size
      $DiscoGb = [math]::Round($Disco.ToInt64($_)/1Gb,0)
      $DiscoLibre = (Get-WmiObject win32_logicalDisk -ComputerName $machine -Filter "DeviceID='C:'").FreeSpace
      $DiscoLibreGb = [math]::Round($DiscoLibre.ToInt64($_)/1Gb,0)
      $RAM = (Get-WMIObject Win32_PhysicalMemory -ComputerName $machine | Measure-Object -Property Capacity -Sum).Sum
      $RAMGb = [math]::Round($RAM.ToInt64($_)/1Gb,2)
      $CPU = (Get-WmiObject -ComputerName $machine Win32_Processor).Name
      $GPU = (Get-WmiObject -ComputerName $machine Win32_VideoController).Name
      $UsersLogged = (Get-CimInstance Win32_LoggedOnUser -ComputerName $machine).Antecedent.name | Select-Object -Unique |Where-Object {$_ -notmatch 'DWM*|UMFD*|SYSTEM|SERVICIO DE RED|SERVICIO LOCAL'}


      New-Object psobject -Property @{
      SystemName = $machine
      IP = $direccionip
      Modelo = $modelo
      Serie = $nserie
      Inventario = $Inventario
      MAC = $MAC
      Velocidad = $interface
      SO = $SO
      LastBootUpTime = $Logindate
      Disco = $DiscoGb
      DiscoLibre = $DiscoLibreGb
      RAM = $RAMGb
      CPU = $CPU
      GPU = $GPU
      Users = $UsersLogged.count,$UsersLogged
      Description = $ADComment

      } | Select-Object SystemName,IP,Modelo,Serie,Inventario,MAC,Velocidad,SO,LastBootUpTime,Disco,DiscoLibre,RAM,CPU,GPU,Users,Description
    }
  else
    {
      $Logindate = (Get-ADComputer $machine -Properties lastlogondate).LastLogonDate
      $SO = (Get-ADComputer $machine -Properties operatingsystem).OperatingSystem

      New-Object psobject -Property @{
      SystemName = $machine
      IP = "Sin Conexion"
      Modelo = ""
      Serie = ""
      Inventario = ""
      MAC = ""
      Velociad = ""
      SO = $SO
      LastBootUpTime = $Logindate
      Disco = ""
      DiscoLibre = ""
      RAM = ""
      CPU = ""
      GPU = ""
      Users = ""
      Description = $ADComment
      } | Select-Object SystemName,IP,Modelo,Serie,Inventario,MAC,Velocidad,SO,LastBootUpTime,Disco,DiscoLibre,RAM,CPU,GPU,Users,Description
    }