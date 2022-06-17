cd e:\[your folder with photo]

$Files = Get-ChildItem -File -Path .

ForEach ( $File in $Files) {
    
    $DateStr = $File.LastWriteTime.ToString("yyyy-MM-dd")

    Write-Host $DateStr 
    
    if ( -not ( Test-Path -Path $DateStr ) ) {

        Write-Host "`n Folder $DateStr does not exist. Creating . . . "

        New-Item -Path . -Name $DateStr -ItemType Directory

    }

    Write-Host $File.fullname

    Move-Item -Path $File.fullname -Dest $DateStr -Force

}
