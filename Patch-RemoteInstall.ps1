###############################################################################################################
#Created by: Cale Pantke
#Modified Date: 12/06/2010 4:00pm
#Modified By: Cale Pantke
###############################################################################################################

#########Remember to edit the variablespath in the main script if you have changed it###########
$psexeclocation = "$PWD" + "\psexec.exe"
set-alias psexec "$psexeclocation"

#Fill out your patches to install. Ex: = "C:\Patching\Patch1.msu", "C:\Patching\Patch2.msu", "C:\Patching\Patch3.msu"
$PatchestoInstall = "$PWD\psexec.exe"
$HotfixID = "KB941158-v2-IE8"
$patchinglog = "$PWD" + "\Logs\patchinglog"
$verifylog = "$PWD" + "\Logs\verifylog"

#Fill out which txt file you need to view by chaning \servers.txt to something like \changed.txt. The -replace removes spaces
#If you need to run servers in a certain order, please remove Sort-Object
$serverstxtpath = "$PWD" + "\servers.txt"
$ServerstoInstall = (get-content "$PWD \servers.txt") -replace ' {1,}','' | Sort-Object

#Fill out which files you would like to send to the machine. This batch file needs to include all patches you want to launch on the machine
$localbatchfile = "$PWD" + "\installpatch.bat"
$batchfile = "installpatch.bat"

#The remote system root drive. This scripts doesn't support multiple system roots yet.
$remotesystemroot = "c:\"

#Name of the directory created on the remote server. This will be cleaned on the clean pass of the main script file
$patchdirectoryname = "Patch\"

#Patch Directory on the target servers
$patchdirectory = "\c$\" + $patchdirectoryname

#Shouldn't need to edit.
$fullpatchpath = "$patchdirectory$batchfile"

MAIN

###############################################################################################################
#Created by: Cale Pantke
#Modified Date: 12/06/2010 4:00pm
#Modified By: Cale Pantke
###############################################################################################################

#####Load the variables script. If it is not there, prompt the user and exit.#####
$PWD = Get-Location
$Variablespath = "$PWD" + "/variables.ps1"

if ((Test-Path -path $Variablespath) -ne $True)
{
Write-Host "EXITING: Please set the correct path in the Variablespath in the main script."
exit
}

. $Variablespath

################################################################################################
################################# VERIFICATION #####################################
################################################################################################

#Find servers.txt in your working directory. Exit if it doesn't exist.
if ((Test-Path -path $serverstxtpath) -ne $True)
{
Write-Host "EXITING: Cannot find servers.txt in your working directory"
exit
}

#Find if you have a Logs directory in your current working directory
$LogDIR = "$PWD" + "\Logs\"
if ((Test-Path -path $LogDIR) -ne $True)
{
Write-Host "EXITING: Cannot find Logs directory in your working directory"
exit
}

#Get Date then set patching log to use current date-time for timestamp
$date = Get-Date
$patchinglog = "$patchinglog{0}{1:d2}{2:d2}-{3:d2}{4:d2}.log" -f $date.year,$date.month,$date.day,$date.hour,$date.minute
$verifylog = "$verifylog{0}{1:d2}{2:d2}-{3:d2}{4:d2}.log" -f $date.year,$date.month,$date.day,$date.hour,$date.minute

################################################################################################
################################# FUNCTIONS ########################################
################################################################################################

Function VerifyUser
{
write-host "You want to install the following patches"
write-host $patchestoinstall
write-host "to the following servers"
write-host $serverstoinstall

$confirm = Read-Host "Is this correct? yes or no."
if ($confirm -ne "yes")
{exit}
else
{Write-Host "You have agreed to install the patches
"}
}

Function CopyFiles
{
#####Create Directory Copy batch file to all servers#####

Foreach ($patch in $patchestoinstall)
{

Foreach ($server in $serverstoinstall)
{
#####Create Patch directory if it isn't there#####
$fullpatchdirectorypath = "\\$server$patchdirectory"
if ((Test-Path -path $fullpatchdirectorypath) -ne $True)
{
Write-Host "Creating Patch directory since it doesn't exist"
New-Item \\$fullpatchdirectorypath -type directory
}

#####Copy all files specified in your patchestoinstall directory #####

Write-Host "
Copying from your current directory " $localbatchfile "to " $fullpatchdirectorypath
Copy-Item $localbatchfile $fullpatchdirectorypath
Write-Host "Copying from your current directory " $patch "to " $fullpatchdirectorypath "since it doesn't exist
"
Copy-Item $patch $fullpatchdirectorypath

}
}
}

Function InstallPatches
{

Foreach ($server in $serverstoinstall)
{
Write-Host "Installing" $patch "to target location" $server
$serverpatchpath = $remotesystemroot + $patchdirectoryname + $batchfile
& psexec -s -i -d \\$server $serverpatchpath

}
}

Function VerifyPatches
{
$confirm = Read-Host "Patching is complete. All machines need to be fully rebooted before starting the Verification stage. Are you ready? yes or no. If you need to skip verification type 'noverify'"
if ($confirm -eq "no")
{
exit
}
elseif ($confirm -eq "noverify")
{
return
}
else
{
Write-Host "Starting verification stage.
"
}

Foreach ($Hotfix in $HotfixID)
{
Foreach ($server in $serverstoinstall)
{
Write-Host "################### Verifing" $hotfix "to target " $server "###################"
Get-WMIObject Win32_QuickFixEngineering -computer $server | where {$_.HotFixID -eq "$Hotfix"}

}
}

}

Function Cleanup
{
Foreach ($server in $serverstoinstall)
{
$fullpatchdirectorypath = "\\$server$patchdirectory"
if ((Test-Path -path $fullpatchdirectorypath) -eq $True)
{
Write-Host "Cleaning up $fullpatchdirectorypath that was created for $server"
Remove-Item \\$fullpatchdirectorypath -recurse
}
if ((Test-Path -path $fullpatchdirectorypath) -ne $True)
{
Write-Host "$fullpatchdirectorypath has been cleaned or were already cleaned from $server"
}
}
}

################################################################################################
################################# MAIN #############################################
################################################################################################

Start-Transcript -path $patchinglog
VerifyUser
CopyFiles
InstallPatches
Stop-Transcript
Start-Transcript -path $verifylog
VerifyPatches
Cleanup
Stop-Transcript
 