$localpath = "$env:ALLUSERSPROFILE\itwonline\"
$pki = "http://web.ad.itwonline.ru/pki/"
$rds = "https://rds.ad.itwonline.ru/rdweb"
$certname = "ITWonline_RCA.cer"
$certfile = "$env:ALLUSERSPROFILE\itwonline\ITWonline_RCA.cer"
$ico = "itwonline.ico"
[String]$StoreName = "root"
[String]$storescope = "CurrentUser"
New-Item $env:ALLUSERSPROFILE\itwonline -itemtype directory 
$client = New-Object System.Net.WebClient
$client.DownloadFileAsync($($pki + $certname), $($localpath + $certname))
Start-Sleep -s 5
$cert = New-Object system.security.cryptography.x509certificates.x509certificate2
$cert.Import($certfile)
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store($StoreName, $StoreScope)
$store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
$store.Add($cert)
$store.Close()
$client.DownloadFileAsync($($pki + $ico), $($localpath + $ico))
Start-Sleep -s 3
$desktop = [Environment]::GetFolderPath("Desktop")
$shell = New-Object -com "WScript.Shell"
$shortcut = $shell.CreateShortcut($desktop + "\ITWonline RDS.lnk")
$shortcut.Arguments = $rds
$shortcut.Description = "Портал веб-приложений ITWonline"
$shortcut.IconLocation = $localpath + $ico
$shortcut.TargetPath = "$env:PROGRAMFILES\Internet Explorer\iexplore.exe"
$shortcut.WindowStyle = 3
$shortcut.Save()
new-object System.Net.WebClient