$ScriptName = "Printers"
#Driver File Name
$FileName = "KX.7.5.0807.zip"
#IPadres of the printer
$PortIP = "10.20.3.204"
$PortName = "IP_$PortIP"
#PrinterName on device
$PrintDriverName = "Kyocera TASKalfa 3051ci KX"
$Printer = "OfficePrinter"

Function Get-ScriptName
{
  split-Path $MyInvocation.ScriptName -Leaf
}
Function RunningAsAdmin {
    $WindowsIdentity=[System.Security.Principal.WindowsIdentity]::GetCurrent()
    $WindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($WindowsIdentity)
    $Admin=[System.Security.Principal.WindowsBuiltInRole]::Administrator
    
    Return $WindowsPrincipal.IsInRole($Admin)
   }
Function InitLogFile {                   
<#---------Logfile Info----------#>
$script:logfile = "c:\temp\$ScriptName-$(get-date -format yyMMddHHmmss).log"
$script:Seperator = @"
$("-" * 25)
"@           
$script:loginitialized = $false           
$script:FileHeader = @"
$seperator
***Application Information***
Filename:  $ScriptName
Created by:  Xander Angenent Avantage
$seperator
"@  
}
Function Write-Log([string]$info){ 
 if($loginitialized -eq $false){ 
  $FileHeader > $logfile                    
  $script:loginitialized = $True
  }                
 $info >> $logfile            
}  
Function Write-CustomOut ($Details){
    #Het doel van deze functie is om een log file te creeren + informatie op scherm te zetten.
	$LogDate = Get-Date -Format T
	Write-Host "$($LogDate) $Details"
	Write-Log "$($LogDate) $Details"
}
Function CleanLogFiles ($CleanDate) {
 $CleanToDay = Get-Date
 $CleanUpDay = $CleanToDay.AddDays(-$CleanDate)
 $CleanLogFiles = Get-ChildItem "$ScriptPad\$ScriptNaamZonderExtentie-*.log"
 Write-CustomOut "Checking for log files older than $CleanUpDay."
 if ($CleanLogFiles) {
  foreach ($CleanLogfile in  $CleanLogFiles) {
   if ( $CleanLogfile.LastWriteTime -le $CleanUpDay) {
    Write-CustomOut "Remove Logfile $($CleanLogFile.Name)"
    Remove-Item $CleanLogfile.Name
   } 
  }
 }  
}
Function Get-ScriptNameZonderExtentie
{
  $TMPScriptNameZonderExtentie = split-Path $MyInvocation.ScriptName -Leaf
  return [io.path]::GetFileNameWithoutExtension($TMPScriptNameZonderExtentie)
}

#Main
InitLogFile
CleanLogFiles 30
Write-CustomOut "Checking if we are running as an Admin user"
If (!(RunningAsAdmin)) {
 Write-CustomOut "Script aborted - We are not admin!"
 exit 1
}
else {
Write-CustomOut "We are admin continue running script."
}
#Create temp directory if it does not exists for logging
if (!(Test-Path("c:\temp"))) { New-Item -Path "c:\" -Name "temp" -ItemType "directory" }
Write-CustomOut "Check if port $PortName exists."
$CheckPortExists = Get-Printerport -Name $PortName -ErrorAction SilentlyContinue
if (!($CheckPortExists)) {
    Write-CustomOut "Creating $PortName"
    Add-PrinterPort -name $PortName -PrinterHostAddress $PortIP
}
#Download driver
Write-CustomOut "Downloading Driver $Filename"
$clnt = new-object System.Net.WebClient
#URL is the URL where the file is.
$url = "https://usa.kyoceradocumentsolutions.com/content/dam/kdc/kdag/downloads/technical/executables/drivers/kyoceradocumentsolutions/us/en/$FileName"
$file = "c:\temp\$FileName"
$clnt.DownloadFile($url, $file)
If (Test-Path($File)) {
  Write-CustomOut "Expanding Driver $Filename"
  Expand-Archive -LiteralPath $file -DestinationPath C:\temp
  #Install Printer
  Write-CustomOut "Installing driver"
  $ReturnInvoke = Invoke-Command {pnputil.exe -a "C:\Temp\kx75_UPD\en\64bit\oemsetup.inf" }
  Write-CustomOut $ReturnInvoke
  Start-Sleep 30
  Write-CustomOut "Adding printer driver."
  Add-PrinterDriver -Name $PrintDriverName
  Start-Sleep 30
  if (Get-PrinterDriver -Name $PrintDriverName) {
    Write-CustomOut "Installing printer $Printer"
    Add-Printer -Name $Printer -PortName $PortName -DriverName $PrintDriverName
}
else {
    Write-CustomOut "Error adding printerdriver!"
}
}
else {
  Write-CustomOut "Download of file $File failed."
}
  Write-CustomOut "Einde"