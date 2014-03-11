Param ( 

    [Parameter( Mandatory = $True, Position = 0, ValueFromPipeline = $False )]
    [String]$User

 )

function find-user {

    $Users = $null
    $Users = Get-ADUser -Filter {SamAccountName -eq $User} -Properties *

    If ( $Users ) {

        Write-Log "User found"
        Return $Users

    } else {

        Write-Log "User not found"

    }

}

function find-profiles {

    $computers = Get-ChildItem \\msk-support\ActiveUsers -Filter ( $ADUser.SamAccountName + '@wts*' ) | %{[regex]::Replace( [regex]::replace( $_,'.*@','' ),'\.txt$','' )}
    return $computers

}

function delete-profiles {

    write-log  ("Начинаем удаление профилей пользователя: " + $ADUser.SamAccountName)
    $computers = @()
    $computers = find-profiles
    write-log "Trying to search user profile on $computers"
    $SAM = $ADUser.SamAccountName

    ForEach ($computer in $computers) {

        $Profile = $null

        Try {

        $Profile = Get-WmiObject Win32_UserProfile -computer $computer -filter "localpath='C:\\Users\\$SAM'"

        } catch {
        #to be continied
        }
        
        If ($Profile) {

            write-log "Profile on $computer is found. Removing"
            $Profile.delete()

        } else {

            write-log "Profile on $computer is not found"

        }
    }
}

function Write-Log {

    Param ( 

        [Parameter( Mandatory = $True, Position = 0, ValueFromPipeline = $True )]
        [String]$Message

     )
        
        Write-Host $Message
        $Message = ( Get-date ).ToString( "T" ) + "  " + $Message
        Out-File $PathToLogFile -InputObject $Message -Append

}

function Get-ScriptDirectory {

    #Данная функция возвращает каталог, в котором хранится наш скрипт
    $Invocation = ( Get-Variable MyInvocation -Scope 1 ).Value
    Split-Path $Invocation.MyCommand.Path
    
}

function Export-MailboxToPST {

    Get-MailboxExportRequest -Status Completed | Remove-MailboxExportRequest -Confirm:$False | Out-Null
    Get-MailboxExportRequest -Status Failed | Remove-MailboxExportRequest -Confirm:$False | Out-Null
    write-log  "Начинаем экспорт почтового ящика $UserMailbox в файл $PSTFilePath"
    $Request = New-MailboxExportRequest -Mailbox $UserMailbox.Alias -FilePath $PSTFilePath
    $Request

}

function Finalize-Script {

    if ( $Session )   {

        Remove-PSSession $Session

    }

    write-log  "Закрываем PSSession до сервера Exchange"
    write-log  "Скрипт завершил работу"
    exit

}

Function Block-User {

    $GroupLogPath = Join-Path ( $CurrentDirectory ) "\$( $ADUser.SamAccountName ).xml"
    ( $ADUser.MemberOf | ConvertTo-XML –NoTypeInformation ).Save( $GroupLogPath )
    write-log "Список групп пользователя выгружен в $GroupLogPath"
    $ADUser | fl | Out-File -FilePath ( [regex]::Replace( $GroupLogPath,'xml$','txt' ) )
    write-log "Выгружены атрибуты учетной записи пользователя в .txt файл"

    foreach ( $Group in $ADUser.MemberOf ) {

        if ( !( $Group -match "DV-users|MS-dax|Domain Users" ) ) {

            write-log  "Пользователь удаляется из группы $Group"
            get-adgroup $Group | Remove-AdGroupMember -member $ADUser.DistinguishedName -Confirm:$false
    
        }
    }
  
    write-log  "Блокируем учетную запись пользователя"
    Disable-ADAccount $ADUser
    write-log  "Удаляем информацию о телефонах"
    $AdUser | Set-ADUser -MobilePhone $Null -OfficePhone $Null -HomePhone $Null
  
    If ( $AdUser.DistinguishedName -match "OU=Sochi" ) {
    
        $MoveToOU = "Sochi"
        
    } else {
    
        $MoveToOU = "Moscow"

    }
  
    $DestinationOU = "OU="+$MoveToOU+",OU=Disabled Users,DC=SOCHI-2014,DC=RU"
    write-log  "Переносим учетную запись в $DestinationOU"
    Move-ADObject $ADUser -TargetPath $DestinationOU
  
}

$CurrentDirectory = Get-ScriptDirectory
import-module activedirectory
$PathToLogFile = $CurrentDirectory + "\" + (Get-Date -format yyyy-MM-dd_HH-mm-ss) + "_" + $User + ".txt"
#Создаем удаленную сессию для импорта командлетов с сервера Exchange
$Session = $null
$ExchServer = 'exch-cas02-n1'
#Пытаемся создать удаленную PS-сессию до сервера
write-log "Попытка создать PSSession до сервера $ExchServer"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ( "http://" +$ExchServer+"/Powershell/" ) -authentication Kerberos -ErrorAction SilentlyContinue

if ( $Session ) {

    write-log  "Удаленная PS-сессия создана";$Session

} else {

    write-log  "Удаленная PS-сессия не создана"
    exit

}

Import-PSSession $Session -AllowClobber | Out-Null

$ADUser = Find-User

if ( $ADUser ){

    $UserMailbox = $Null
    if ( $ADUser.mail ) {
    
        $UserMailbox = get-mailbox $( $ADUser.mail )
        
        if ( $UserMailbox ) {
     
            $PSTFilePath = "\\exch-mb27\e$\PST\"+$ADUser.sAMAccountName+".pst"
            $ExportRequest = Export-MailboxToPST
            $LastRequestStatus = $Null   

            do {

                $ExportRequestStatus = ( Get-MailboxExportRequest $ExportRequest.RequestGuid ).Status

                if ( $LastRequestStatus -ne $ExportRequestStatus ) {

                    $LastRequestStatus = $ExportRequestStatus
                    $OutString = "Статус запроса на экспорт почтового ящика изменен на "+$LastRequestStatus
                    write-log $OutString

                }

                Start-Sleep 15
       
            } until ( ( $ExportRequestStatus -eq "Completed" ) -or ( $ExportRequestStatus -eq "Failed" ) )
    
            if ( $ExportRequestStatus -eq "Failed" ) {

                write-log "Экспорт почтового ящика в PST завершился ошибкой"

            } elseif ( ( Get-Item $PSTFilePath ).Length -eq 271360 ) {

                write-log "Экспорт почтового ящика в PST завершился ошибкой. Размер выходного файла равен 265Kb. Необходимо предоставить учетной записи, от имени которой запускается скрипт, доступ уровня Full Access к данному почтовому ящику и запустить скрипт повторно."
    
            } else {
    
                $NewPSTFilePath = "\\file-06\d$\PST_backup\Удаленные с сервера\" + $ADUser.sAMAccountName + ".pst"
                write-log "Переносим выгруженный файл в $NewPSTFilePath"
                move-item -LiteralPath $PSTFilePath -Destination $NewPSTFilePath -Force
                write-log "Отключаем почтовый ящик"
                Disable-Mailbox $UserMailbox.Alias -Confirm:$False
                Block-User

            }

        } else {

            write-log "У пользователя нет почтового ящика"
            delete-profiles
            Block-User
        
        }

    } else {

        write-log "У пользователя нет почтового ящика"
        delete-profiles
        Block-User

    }    
}

Finalize-Script