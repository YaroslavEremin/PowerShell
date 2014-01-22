
function Set-PicToADUserScheduledTask {
<#
  .SYNOPSIS
  This function set pictures for Exchange user
  .DESCRIPTION
  The function is designed to run as a scheduled task.
  .EXAMPLE
  .\Set-PicToADUserScheduledTask.ps1 -ExchSrv exch-cas01-n1 -PicsDir \\exch-cas01-n1\c$\USER_PICS -LogFileName _LOG.csv
  .EXAMPLE
  .\Set-PicToADUserScheduledTask.ps1 -PicsDir \\exch-cas01-n1\c$\USER_PICS\ -LogFileName _LOG.csv
  .PARAMETER ExchSrv
  Set target Exchange mailbox server
  .PARAMETER $PicsDir
  Set target folder with user pictures
  .PARAMETER $LogFileName
  Set name for log file
  #>
param (
[Parameter(Mandatory=$False)] $ExchSrv = $env:COMPUTERNAME,
[Parameter(Mandatory=$True)] $PicsDir,
[Parameter(Mandatory=$True)] $LogFileName
)
$Session = New-PSSession -ConnectionUri ("http://" + $ExchSrv +"/Powershell/") -Authentication Kerberos -ConfigurationName Microsoft.Exchange -ErrorAction SilentlyContinue
Import-PSSession $Session -AllowClobber | Out-Null
If ( $PicsDir -match "\$" ) {
$PicsDir = [regex]::replace($PicsDir,'\$','')
}
$LogFilePath = $PicsDir + "\" + $LogFileName
$ArchivePicDir = $PicsDir + "\" + "Archive"
$NewPics = Get-ChildItem -Path $PicsDir -Filter "*.jpg"
$NewPics = Sort-Object -InputObject $NewPics -Property LastWriteTime
#Checking for pictures folder
If ( -not(Test-Path $PicsDir) ) {
New-Item $PicsDir -ItemType Directory
}
#Checking for pictures archive folder
If ( -not(Test-Path $ArchivePicDir) ) {
New-Item $ArchivePicDir -ItemType Directory
}
#Checking for log file
If (-not(Test-Path $LogFilePath) ) {
New-Item ($PicsDir + "\" + $LogFileName) -ItemType File 
$LogObject = Get-Item -Path ($PicsDir + "\" + $LogFileName) -Include $LogFileName
$LogObject = Select-Object -InputObject $LogObject -Property Name,LastWriteTime
Export-Csv -InputObject $LogObject -Path $LogFilePath -Append
}
#Do job
ForEach ($NewPic in $NewPics) {
$UserName = [regex]::replace($NewPic.Name,'.jpg$','')
$path = $PicsDir + '\' + $NewPic.Name
#Set pictures for user
Import-RecipientDataProperty -Identity $UserName -Picture -FileData ([Byte[]]$(Get-Content -path $path -Encoding Byte -ReadCount 0))
$LogObject = Select-Object -InputObject $NewPic -Property Name,LastWriteTime
#Record job to a log file
$NewPics = Sort-Object -InputObject $NewPics -Property LastWriteTime
#Checking for pictures folder
If ( -not(Test-Path $PicsDir) ) {
New-Item $PicsDir -ItemType Directory
}
#Checking for pictures archive folder
If ( -not(Test-Path $ArchivePicDir) ) {
New-Item $ArchivePicDir -ItemType Directory
}
#Checking for log file
If (-not(Test-Path $LogFilePath) ) {
New-Item ($PicsDir + "\" + $LogFileName) -ItemType File 
$LogObject = Get-Item -Path ($PicsDir + "\" + $LogFileName) -Include $LogFileName
$LogObject = Select-Object -InputObject $LogObject -Property Name,LastWriteTime
Export-Csv -InputObject $LogObject -Path $LogFilePath -Append
}
#Do job
ForEach ($NewPic in $NewPics) {
$UserName = [regex]::replace($NewPic.Name,'.jpg$','')
$path = $PicsDir + '\' + $NewPic.Name
#Set pictures for user
#Import-RecipientDataProperty -Identity $UserName -Picture -FileData ([Byte[]]$(Get-Content -path $path -Encoding Byte -ReadCount 0))
$LogObject = Select-Object -InputObject $NewPic -Property Name,LastWriteTime
#Record job to a log file
Format-Table -InputObject $LogObject -AutoSize | Out-File -FilePath $LogFilePath -Append
Move-Item -Path $path -Destination ($ArchivePicDir + "\" + $NewPic.Name) -Force
}
Remove-PSSession $Session
}
}
