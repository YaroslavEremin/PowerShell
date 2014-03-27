Param ( 

    [switch]$debugMode = $false

 )

 # Attempts to open a file and trap the resulting error if the file is already open/locked
 function Test-FileLock {
    
    param (
    
        [string]$filePath
        
    )

    $filelocked = $true
    $fileInfo = New-Object System.IO.FileInfo $filePath

    trap {

        Set-Variable -name Filelocked -value $false -scope 1
        continue

    }

    $fileStream = $fileInfo.Open( [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None )
    
    if ($fileStream) {

        $fileStream.Close()

    }

    $filelocked

}

#Поиск пользователей в целевых OU, отключенных заданное количество дней назад
function find-users {

    $Date = (Get-Date)
    $Date = $Date.AddDays(-3)
    $Users = $NULL

    Try {

        $Users = @( Get-QADUser -LastChangedBefore $Date -Disabled -SearchRoot "OU=Sochi,DC=SOCHI-2014,DC=RU" )
        $Users += Get-QADUser -SearchRoot "OU=Sochi,DC=SOCHI-2014,DC=RU" -AccountExpiresBefore $Date

    } Catch {

        write-log $Error[0].Exception
        write-log $Error[0].InvocationInfo.Line

    }

    Try {

       $Users += Get-QADUser -LastChangedBefore $Date -Disabled -SearchRoot "OU=Moscow,DC=SOCHI-2014,DC=RU"
       $Users += Get-QADUser -SearchRoot "OU=Moscow,DC=SOCHI-2014,DC=RU" -AccountExpiresBefore $Date

    } Catch {

        write-log $Error[0].Exception
        write-log $Error[0].InvocationInfo.Line

    }

    If ( $Users ) {

        Write-Log "Users found"
        $Users = $Users | Select-Object -First 20
        Return $Users

    } else {

        Write-Log "Users not found"
        Exit

    }

}

#Поиск на какие компьютеры заходил пользователь
function find-profiles {

    $computers = Get-ChildItem \\msk-support\ActiveUsers -Filter ( $ADUser.SamAccountName + '@wts*' ) | %{[regex]::Replace( [regex]::replace( $_,'.*@','' ),'\.txt$','' )}
    return $computers

}

#Удаление профилей пользователя на найденных компьютерах
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
        
             write-log $Error[0].Exception
             write-log $Error[0].InvocationInfo.Line

        }
        
        If ($Profile) {

            write-log "Profile on $computer is found. Removing"

            If ( -not($debugMode) ) {

                Try {

                    $Profile.delete()

                } Catch {

                    write-log $Error[0].Exception
                    write-log $Error[0].InvocationInfo.Line

                }

            }

        } else {

            write-log "Profile on $computer is not found"

        }

    }

}

#Логирование действий
function Write-Log {

    Param ( 

        [Parameter( Mandatory = $True, Position = 0, ValueFromPipeline = $True )]
        [String]$Message

     )
        
        Write-Host $Message
        $Message = ( Get-date ).ToString( "T" ) + "  " + $Message
        Out-File $PathToLogFile -InputObject $Message -Append

}

#Данная функция возвращает каталог, в котором хранится наш скрипт
function Get-ScriptDirectory {

    $Invocation = ( Get-Variable MyInvocation -Scope 1 ).Value
    Split-Path $Invocation.MyCommand.Path
    
}

#Запуск экспорта почтовых ящиков отключенных пользователей
function Export-MailboxesToPST {

    Get-MailboxExportRequest -Status Failed | Remove-MailboxExportRequest -Confirm:$False -ErrorAction SilentlyContinue | Out-Null


    ForEach ($ADUser in $ADUsers) {

        $UserMailbox = $Null

        If ( $ADUser.mail ) {

            Try {

                $UserMailbox = get-mailbox $( $ADUser.mail )

            } Catch {

                write-log $Error[0].Exception
                write-log $Error[0].InvocationInfo.Line

            }

        }

        if ( $UserMailbox ) {
              
            if ( !(Get-MailboxExportRequest -Mailbox $UserMailbox.Alias ) ) {

                write-log  ( "Начинаем экспорт почтового ящика " + $ADUser.SamAccountName + " в файл." )
            
                Try {

                    New-MailboxExportRequest -Mailbox $UserMailbox.Alias -FilePath ( "\\exch-mb27\e$\PST\" + $ADUser.sAMAccountName + ".pst" ) -WhatIf:$debugMode
            
                } catch {

                    write-log $Error[0].Exception
                    write-log $Error[0].InvocationInfo.Line
         
                }

            }
        } else {

            write-log ( "У пользователя " + $ADUser.SamAccountName + " нет почтового ящика" )
            Set-ADUser -Identity $ADUser.SamAccountName -EmailAddress $NULL
            
        }

    }

}

function Restart-FailedMailboxExportRequests {

    $FaiedExportRequests = Get-MailboxExportRequest -Status Failed

    ForEach ( $FaiedExportRequest in $FaiedExportRequests ) {

        write-log ( "Экспорт почтового ящика " + $FaiedExportRequest.Mailbox + " в PST завершился ошибкой" )
        $FaiedExportRequest | Resume-MailboxExportRequest -WhatIf:$debugMode
        write-log  "Возобновляем экспорт почтового ящика в PST"

    }

}

function Restart-MailboxExportRequestsWithFullRights {

    $CompletedExportRequests = MailboxExportRequest -Status Completed
                   
    ForEach ( $CompletedExportRequest in $CompletedExportRequests ) {

        If ( ( Get-Item $CompletedExportRequest.FilePath ).Length -eq 271360 ) {

            write-log ( "Экспорт почтового ящика " + $CompletedExportRequest.Mailbox + " в PST завершился ошибкой. Размер выходного файла равен 265Kb. Необходимо предоставить учетной записи, от имени которой запускается скрипт, доступ уровня Full Access к данному почтовому ящику." )
            $CurrentUser = [Environment]::UserName
            Add-MailboxPermission -Identity $CompletedExportRequest.Mailbox -User $CurrentUser -AccessRights 'FullAccess' -ErrorAction SilentlyContinue | Out-Null
            Start-Sleep 15
            write-log ('Удаляем файл экспорта ' + $CompletedExportRequest.FilePath)
            Remove-Item -Path $CompletedExportRequest.FilePath -Force -Confirm:$False -WhatIf:$debugMode -ErrorAction SilentlyContinue
            write-log "Удаляем запрос экспорта"
            $CompletedExportRequest | Remove-MailboxExportRequest -WhatIf:$debugMode -Confirm:$False
            write-log "Пересоздаем запрос экспорта"
            New-MailboxExportRequest -Mailbox $CompletedExportRequest.Mailbox -FilePath $CompletedExportRequest.FilePath  -WhatIf:$debugMode -Confirm:$False
            
        }

    }

}

function Disable-MailboxWithComplitedExport {

    $CompletedExportRequests = Get-MailboxExportRequest  | ? { $_.Status -match "Completed" }

    ForEach ( $CompletedExportRequest in $CompletedExportRequests ) {

        write-log ( "Экспорт почтового ящика " + $CompletedExportRequest.Mailbox + " в PST завершен." )
        $CompletedExportRequest | Remove-MailboxExportRequest -WhatIf:$debugMode -Confirm:$False
        write-log "Отключаем почтовый ящик"
        Disable-Mailbox $CompletedExportRequest.Mailbox -Confirm:$False -WhatIf:$debugMode

    }

}

function Check-ExportFinishing {

    $Flag = $False
    $ExportRequests = Get-MailboxExportRequest
    $FaiedExportRequests = $ExportRequests | ? { $_.Status -eq "Failed" }
        
    If ( -not($ExportRequests) ) {

        $Flag = $True
        write-log "Все экспорты завершены"

    } ElseIf ( $ExportRequests.Count -eq $FaiedExportRequests.Count ) {

        $Flag = $True
        $FaiedExportRequests | Remove-MailboxExportRequest -WhatIf:$debugMode -Confirm:$False

        ForEach ($FaiedExportRequest in $FaiedExportRequests) {

            write-log ( "Экспорт почтового ящика " + $FaiedExportRequest.Mailbox + " в PST завершился ошибкой" )
               
        }

    }

    Return $Flag
}

function Check-ExportState {

    $Timer = 90

    do {
        
        Write-Host ( ( Get-date ).ToString( "T" ) + " Waiting of export completion" ) -ForegroundColor Green

        Try {

            Get-MailboxExportRequest | Get-MailboxExportRequestStatistics | Select-Object -Property SourceAlias,SourceDatabase,PercentComplete,BytesTransferredPerMinute | Sort-Object -Property Status,Identity | Format-Table -AutoSize
                    
        } Catch {

                write-log $Error[0].Exception
                write-log $Error[0].InvocationInfo.Line

        }
        
        Start-Sleep $Timer
        $Timer = $Timer - 1

                
        Restart-FailedMailboxExportRequests
        Restart-MailboxExportRequestsWithFullRights
        Disable-MailboxWithComplitedExport
        Check-ExportFinishing

    } until ( Check-ExportFinishing -or $Timer -eq 0)

}

#Перемещение выгруженных в PST файлы почтовых ящиков на файловый ресурс архива.
function Move-ArchivePSTFiles {

    ForEach ( $ADUser in $ADUsers ) {

        $PSTFilePath = "\\exch-mb27\e$\PST\" + $ADUser.sAMAccountName + ".pst"
        $NewPSTFilePath = "\\file-06\d$\PST_backup\Удаленные с сервера\" + ( [regex]::match( $PSTFilePath, "[^\\]*$" ).Value )

        If ( ( Test-Path $PSTFilePath ) -and ( Test-FileLock $PSTFilePath ) ) {
    
            write-log ( "Переносим файл экспорта в новое расположение " + $NewPSTFilePath )
            move-item -LiteralPath $PSTFilePath -Destination $NewPSTFilePath -Force -Confirm:$False -WhatIf:$debugMode -ErrorAction SilentlyContinue
            
        }
                
    }

}

function Clean-ExportDirectory {

write-log  'Переносим все файлы из \\exch-mb27\e$\PST\ в \\file-06\d$\PST_backup\Удаленные с сервера\'
move-item -Path "\\exch-mb27\e$\PST\*.pst"  -Destination "\\exch-mb27\e$\PST" -Force -Confirm:$False -WhatIf:$debugMode -ErrorAction SilentlyContinue
            
}

#Закрытие удаленной сессии PS и завершение работы скрипта.
function Finalize-Script {

    if ( $Session ) {

        Remove-PSSession $Session

    }

    write-log  "Закрываем PSSession до сервера Exchange"
    write-log  "Скрипт завершил работу"
    exit

}

function Clean-ADUserGroups {

    foreach ( $Group in $ADUser.MemberOf ) {

        if ( !( $Group -match "DV-users|MS-dax|Domain Users" ) ) {

            write-log ( "Пользователь " + $ADUser.SamAccountName + " удаляется из группы $Group")

            Try {

                Remove-AdGroupMember -Identity $Group -members $ADUser.SamAccountName -Confirm:$false -WhatIf:$debugMode
    
            } catch {

                write-log $Error[0].Exception
                write-log $Error[0].InvocationInfo.Line

            }

        }

    }

}

function Move-ADUserToOUForDisabledUsers {

    If ( $AdUser.DN -match "OU=Sochi" ) {
    
        $MoveToOU = "Sochi"
        
    } else {
    
        $MoveToOU = "Moscow"

    }
  
    $DestinationOU = "OU="+$MoveToOU+",OU=Disabled Users,DC=SOCHI-2014,DC=RU"
    write-log  "Переносим учетную запись $Username в $DestinationOU"

    Try {

        Get-ADUser -Identity $ADUser.SamAccountName | Move-ADObject -TargetPath $DestinationOU -WhatIf:$debugMode

    } catch {

        write-log $Error[0].Exception
        write-log $Error[0].InvocationInfo.Line

    }

}

Function Block-User {

    $GroupLogPath = Join-Path ( $CurrentDirectory ) "\$( $ADUser.SamAccountName ).xml"
    ( $ADUser.MemberOf | ConvertTo-XML –NoTypeInformation ).Save( $GroupLogPath )
    write-log ( "Список групп пользователя " + $ADUser.SamAccountName + " выгружен в $GroupLogPath" )
    $ADUser | fl | Out-File -FilePath ( [regex]::Replace( $GroupLogPath,'xml$','txt' ) )
    write-log "Выгружены атрибуты учетной записи пользователя $Username в .txt файл"
    Clean-ADUserGroups
    write-log ( "Блокируем учетную запись пользователя " + $ADUser.SamAccountName )
    Disable-ADAccount $ADUser.SamAccountName -WhatIf:$debugMode
    write-log ( "Удаляем информацию о телефонах пользователя " + $ADUser.SamAccountName )
    Set-ADUser -Identity $AdUser.SamAccountName -MobilePhone $Null -OfficePhone $Null -HomePhone $Null -WhatIf:$debugMode
    Move-ADUserToOUForDisabledUsers

  
}

function Create-PSSessionToExchange {

    param (

        [string]$ExchServer = 'exch-cas02-n1'

    )

    write-log "Попытка создать PSSession до сервера $ExchServer"
    $ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ( "http://" + $ExchServer + "/Powershell/" ) -authentication Kerberos -ErrorAction SilentlyContinue
    Import-PSSession $ExchSession -AllowClobber | Out-Null
    if ( $ExchSession ) {

        write-log  "Удаленная PS-сессия создана";
        Return $ExchSession

    } else {

        write-log  "Удаленная PS-сессия не создана"
        exit

    }




}

$CurrentDirectory = Get-ScriptDirectory
$ErrorActionPreferenceState = $ErrorActionPreference
$ErrorActionPreference = "Stop"
import-module activedirectory
add-pssnapin quest.activeroles.admanagement
$PathToLogFile = $CurrentDirectory + "\Clear-DisabledUsers_" + (Get-Date -format yyyy-MM-dd_HH-mm-ss) + ".txt"
$Session = $null
$Session = Create-PSSessionToExchange
$ADUsers = $NULL
$ADUsers = Find-Users
$ADUsers | ft -AutoSize SamAccountName,Description,Title
Start-Sleep 10
Write-Host  "Export Mailboxes To PST" -ForegroundColor Green
#Clean-ExportDirectory
Export-MailboxesToPST
Write-Host "Check Export State" -ForegroundColor Green
Check-ExportState
Start-Sleep 15
Write-Host "Move Archive PST Files" -ForegroundColor Green
#Move-ArchivePSTFiles

ForEach ( $ADUser in $ADUsers ) {

    delete-profiles
    Block-User
        
}

$ErrorActionPreference = $ErrorActionPreferenceState
Finalize-Script