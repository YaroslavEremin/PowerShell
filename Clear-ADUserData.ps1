Param ( 

    [Parameter( Mandatory = $True, Position = 0, ValueFromPipeline = $False )]
    [String]$User

 )

function find-user {

    $Users = $null
    $Users = Get-ADUser -Filter {SamAccountName -eq $User} -Properties *

    If ( $Users ) {

        Write-Log "User found"

    } else {

        Write-Log "User not found"

    }

}

function find-profiles {

    Param ( 

        [Parameter( Mandatory = $True, Position = 0, ValueFromPipeline = $True )]
        [String]$SAM

     )

    $computers = Get-ChildItem \\msk-support\ActiveUsers -Filter ( $SAM + '@*' ) | %{[regex]::Replace( [regex]::replace( $_,'.*@','' ),'\.txt$','' )}
    return $computers

}

function delete-profiles {

    Param ( 

        [Parameter( Mandatory = $True, Position = 0, ValueFromPipeline = $False )]
        [String]$SAM,
        [Parameter( Mandatory = $True, Position = 1, ValueFromPipeline = $True )]
        [String]$computer

     )

    write-log  "Начинаем удаление профилей пользователя"

    If ( Test-Connection $computer ) {

        $PathEN = "\\" + $computer + "\c`$\Users\" + $SAM
        $PathRU = "\\" + $computer +  "\c`$\Пользователи\" + $SAM

        If ( Test-Path $PathEN ) {

            Remove-Item $PathEN -Recurse -Force -Confirm:$False -ErrorAction SilentlyContinue -whatif:$debugMode
            Write-Log "Delete profile: $PathEN"

        } ElseIf ( Test-Path $PathRU ) {

            Remove-Item $PathRU -Recurse -Force -Confirm:$False -ErrorAction SilentlyContinue -whatif:$debugMode
            Write-Log "Delete profile: $PathRU"

        } Else {

            Write-Log "Profile not find"

        }

    } Else {

        Write-Log "Computer is offline"

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
    $Request = New-MailboxExportRequest -Mailbox $UserMailbox.Alias -FilePath $PSTFilePath -whatif:$debugMode
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
            get-adgroup $Group | Remove-AdGroupMember -member $ADUser.DistinguishedName -Confirm:$false -whatif:$debugMode
    
        }
    }
  
    write-log  "Блокируем учетную запись пользователя"
    Disable-ADAccount $ADUser -whatif:$debugMode
  
    $Phones = @{
        "HomePhone"       = $Null
        "mobile"          = $Null
        "OfficePhone"     = $Null 
    }
  
    write-log  "Удаляем информацию о телефонах"
    $AdUser | Set-ADUser @Phones -whatif:$debugMode
  
    If ( $AdUser.DistinguishedName -match "OU=Sochi" ) {
    
        $MoveToOU = "Sochi"
        
    } else {
    
        $MoveToOU = "Moscow"

    }
  
    $DestinationOU = "OU="+$MoveToOU+",OU=Disabled Users,DC=SOCHI-2014,DC=RU"
    write-log  "Переносим учетную запись в $DestinationOU"
    Move-ADObject $ADUser -TargetPath $DestinationOU -whatif:$debugMode
  
}

$CurrentDirectory = Get-ScriptDirectory
import-module activedirectory
$PathToLogFile = $CurrentDirectory + "\" + (Get-Date -format yyyy-MM-dd_HH-mm-ss) + "_" + $User
#Создаем удаленную сессию для импорта командлетов с сервера Exchange
$Session = $null
$ExchServer = 'exch-cas02-n1'
#Проходим по списку и пытаемся создать удаленную PS-сессию до сервера
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
    
            $NewPSTFilePath = "\\file-06\d$\PST_backup\Удаленные с сервера\"+$ADUser.sAMAccountName+".pst"
            write-log "Переносим выгруженный файл в $NewPSTFilePath"
            move-item $PSTFilePath $NewPSTFilePath -whatif:$debugMode -Force
            write-log "Отключаем почтовый ящик"
            Disable-Mailbox $UserMailbox.Alias -Confirm:$False -whatif:$debugMode
            Block-User

        }
    
        } else {

        write-log "У пользователя нет почтового ящика"
        Block-User

        }

    } else {

    write-log "У пользователя нет почтового ящика"
    Block-User

    }    
}

Finalize-Script