#This script
# a) creates VACATION mailbox event with Out-of-Office time mark
# b) turn OOF on if necessary.
#
# Verison 4.0
# (c) Aleksandr Ilyushenko, 2022
#
[Reflection.Assembly]::LoadFile("C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll") | Out-Null
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
#Init data. Gets from init parameters' file
    $PS_Script_Tilte = ""
    $log_msg = @{}
    $c_in =""
    $domain_rb = ""
    $domain_gts = ""
    $domain_gts_SearchBase = ""
    $connectionString = ""
    $EWSURL = '' #
    $logPath = $env:USERPROFILE +"\Documents\add_OOF_psscript\logs\"
    
    
    #Send-myMessage constants
    $myFrom = ""
    $mySubj = ""
    $mySmtpServer = ""
    $myTo = ""

#load init parameters' file
    ."C:\Scripts\al_PSScripts\OOF_init.ps1"
    

Function Send-alMessage {
    param (
        $body,
        $To,
        $Subj
    )
    $my_error = ""
    $body += $body_info
    Send-MailMessage -From $myFrom -Subject $Subj -SmtpServer $mySmtpServer -Body $body -To $To -Encoding unicode -BodyAsHtml -ErrorVariable my_error -ErrorAction SilentlyContinue
    if ($my_error -ne "") {
        $body =  "Error : " + $my_error[0].Exception.Message
        $body += "<br>" + ($my_error[0].CategoryInfo -join '`r`n')
        $body += "<br>" + $my_error[0].ScriptStackTrace
        Write-alError -EntryType Warning -Message $body
    }
}

Function Write-alError {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $EntryType,
        [Parameter(Mandatory=$true, Position=1)]
        [string] $Message
    )
    #add the script path/file-name
    $Message = $PSCommandPath + " : PS-Script name" + "`r`n`r`n" + $Message
    Write-EventLog -LogName "Application" -Source "PSScriptOOF" -EventID 2 -EntryType $EntryType -Message $Message -Category 1 -RawData 10,20
}

Function Format-alInfo {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $init_str,
        [Parameter(Mandatory=$true, Position=1)]
        [string] $target_str
    )
    
    $search_str = $init_str + '.*?.*'
    $init_str = $init_str + ": "
    $final_str = $target_str + "`r`n" + $init_str + $now_time
    if ($target_str  -like "*$init_str*") {
        $info = $target_str -split "`r?`n"
        #$info | %{$_ -like 'Last c*'}
        $final_str = ($info | %{$_ -replace $search_str, ($init_str + $now_time)}) -join "`r`n"
    }
    $final_str
}

Function Format-alCountWorkingDays {
#get j$ working days before vacation start day

    param (
        [Parameter(Mandatory=$true, Position=0)]
        [int] $j
    )
        #$j += 1
        for ($i = 0; $i -gt $j-1; $i--) {
            $tmp1 = Get-date($d_start).AddDays($i)
            if ($tmp1.DayOfWeek -eq 'Sunday') {$i-=2; $j += -2}
            if ($tmp1.DayOfWeek -eq 'Saturday') {$i-=1; $j += -1}
            #$tmp1 = Get-date($d_start).AddDays($i)
        }
        return Get-date($d_start).AddDays($j)
        
}

Function Remove-alLogFiles {
    param (
        $log_Path,
        $logDepth
        )
    $logList = (Get-ChildItem -Path $log_Path).Name
    $nowt = Get-Date
    foreach ($logf in $logList) {
        try {
            (($logf | %{($_ -split "\.")[0]} | %{[datetime]::parseexact($_, 'yyyyMMdd_HHmmss', $null)}) | %{$_ - $nowt}).days | %{ if ($_ -lt -$logDepth) {Remove-Item -Path ($log_Path + $logf)}}
        } catch {
            $logf
        }
    }
    #(($logList | %{($_ -split "\.")[0]} | %{[datetime]::parseexact($_, 'yyyyMMdd_HHmmss', $null)}) | %{$_ - $nowt}).days | %{ if ($_ -lt -2) {$_}}
    
}


try {
  #Write-alError -EntryType Information -Message ("Start")
#Event log readyness check 
    if (((Get-WmiObject -Class Win32_NTEventLOgFile |
        Select-Object FileName, Sources |
        ForEach-Object -Begin { $hash = @{}} -Process { $hash[$_.FileName] = $_.Sources } -end { $Hash })["Application"] -match "PSScriptOOF").Count -eq 0) {
        New-EventLog -LogName Application -Source "PSScriptOOF"
        Start-Sleep 3
    }

#get current vacation data
  
    $now_date = (Get-date).Date
    $body1 = ""
    $bodyTmp = ""
    $file = $env:USERPROFILE +"\Documents\add_OOF_psscript\vacation_list.csv"
    $file_log = $logPath + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt"
    #Write-alError -EntryType Information -Message ($file)

    $tmp2 = ""
    $param = @()
    if (Test-Path -Path $file -PathType Leaf) {
        $tmp2 = Import-Csv -Path $file -Delimiter ";" -Encoding UTF8 #|select SAMAccountNameAD,PregLeaveName,PregLeaveFrom,PregLeaveTo, @{n='StartShift'; e={0}}, @{n='EndShift'; e={0}}, @{n='IsEventExist'; e={$false}}, @{n='IsOOFOn'; e={$false}}, @{n='FullString'; e={$_.SAMAccountNameAD + $_.PregLeaveName + $_.PregLeaveFrom + $_.PregLeaveTo}}
        $param = $tmp2[0].psobject.Properties.Name
        }

    
    $vacRecActual = $tmp2 | ?{((Get-Date($_.PregLeaveTo) -ErrorAction SilentlyContinue).Date) -ge $now_date}
    
    $body1 += "<br>" + $vacRecActual.Count + " : Count of actual vacation records"
   
#get email's list
    

#get email's list at users' domain
    
    $memb_new_rosbank_email = ""
    #Write-alError -EntryType Information -Message ("select")
    $memb_new_rosbank_email = $vacRecActual | select *, @{n='mail';e={%{(Get-ADUser -Server $domain_rb -LDAPFilter "(&(SamAccountName=$($_.SAMAccountNameAD))(!userAccountControl:1.2.840.113556.1.4.803:=2))" -Properties mail -ErrorAction SilentlyContinue | select mail).mail}}}
    
    #get existed email
    $email_exist = $memb_new_rosbank_email | ?{$_.mail -ne $null}
    #$email_exist.Count
    
    #get existed mailboxes
    $mailbox_exist = $memb_new_rosbank_email | select *, @{n='isMailBox';e={ if($_.mail -ne $null) {if((get-mailbox $_.mail -ErrorAction SilentlyContinue) -ne $null) {1}}}}

    #get mailboxes to proceed
    $mailBoxes = $mailbox_exist | ?{$_.isMailBox -eq 1}
    $body1 += "<br>" + $mailBoxes.Count + " : Count of existed maiboxes"
    $mailBoxes_not = $mailbox_exist | ?{(-not($_.isMailBox -eq 1))}
    #$mailBoxes.Count
    #$mailBoxes_not.Count

    #Write-alError -EntryType Information -Message ("body")
    $body = ""
    #$body1 = ""

    #$mailBoxes_temp = $mailBoxes |select -First 1
    
    $logFile = "Log_flags;" + ($param -join ";") + ";mail;isMailBox`r`n"
    $logFileErrors = $false

    $MustEventCount = 0 #Count of events should be exist
    $CreatedEventCount = 0 #Count of events to be created
    $NoEventCount = 0 #Count of events that are not exist
    $NoNeedEventCount = 0 #Count of events that are not to be necessary created
    $OOF_OffCount = 0 #Count of OOF-Off records should be checked
    $OOF_OnCount = 0 #Count of OOF turned On
    $OOF_OnMessageSentCount = 0 #Count of 'OOF turned On' messages have been sent to user 

    $TZ_default = [System.TimeZoneInfo](Get-TimeZone)
    $userAppointments =@()
    $EWSTZ = @()
    $vacRecord = @()
    $recErrors = @()

    #(0..($mailboxes.Count-1)) | where {$mailboxes[$_].SAMAccountNameAD -eq ''}


#$i = 2307 20 4
    for ($i = 0; $i -lt $mailBoxes.Count; $i++) {
         #Write-alError -EntryType Information -Message ($mailBox)
        
        $d_start = Get-Date($mailBoxes[$i].PregLeaveFrom) # -Format 'dd.MM.yyyy'
        $d_end = Get-Date($mailBoxes[$i].PregLeaveTo) # -Format 'dd.MM.yyyy'
        
        #$MustEventCount += 1
        #$mailBoxes[$i].IsEventExist = $false
        #$mailBoxes[$i].IsOOFOn = $false
        
        #Exception calling "FindAppointments" with "1" argument(s): "The specified view range exceeds the maximum range of two years."
        if (($d_end - $d_start).Days -gt 600) { $mailBoxes[$i].IsEventExist = $true;  $MustEventCount -= 1 } # elimination of  '2 years over' error 

        $nowTimeDelta = ($d_start.AddDays(-5) - $now_date).Days
        if ($nowTimeDelta -le 0) { $MustEventCount += 1; $logFile += "EventMustExist" }
        
        #vacation event existance control
        if (($nowTimeDelta -le 0) -and ($mailBoxes[$i].IsEventExist -ne $true)) {

        $logFile += ",5,EventNotCreated"
        
        $appStart = $d_start.ToString('dd.MM')
        $appEnd = $d_end.ToString('dd.MM')
        $oof_def_text = "Уважаемые коллеги, добрый день. С "  + $appStart + " по " + $appEnd + "  нахожусь в отпуске."

        #check vacation event existance
        
        $appointments = $null
        
       
        $EmailAddress = $mailBoxes[$i].mail
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
        
        try {
            $tz1 = Get-TimeZone -id (Get-MailboxRegionalConfiguration $EmailAddress).TimeZone
            $EWSTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById($tz1.id)
        } catch {
            $EWSTimeZone = $TZ_default
        }
       
        $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion,$EWSTimeZone)
        $Service.Url = [System.URI] $EWSURL
        $service.UseDefaultCredentials = $true
        
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress)
        $CalendarFolder = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)

        #Exception calling "FindAppointments" with "1" argument(s): "You have exceeded the maximum number of objects that can be returned for the find operation. Use paging to reduce the result size and try your request again"
        #use view upper limit 999
        $View = [Microsoft.Exchange.WebServices.Data.CalendarView]::new($d_start.AddSeconds(1),$d_end.AddDays(1),999)
        $View.PropertySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject,
                             [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start,[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End,
                             [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::DateTimeSent,
                             [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::LegacyFreeBusyStatus,
                             [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsAllDayEvent,
                             [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::TimeZone)
          #                   [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::StartTimeZone,
          #                   [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::EndTimeZone)
        
        #$appointments = $CalendarFolder.FindAppointments($View) | ?{$_.Subject -like 'Отпуск*(автоматическое событие)'}# |  Select-Object start , End , Subject 
        #$appointments = $CalendarFolder.FindAppointments($View) | ?{($_.LegacyFreeBusyStatus -eq 'OOF') -and ($_.IsAllDayEvent -eq $true) }# |  Select-Object start , End , Subject 
        #Write-alError -EntryType Information -Message ("17")
            try {  
                $CalendarFolder.FindAppointments($View) | ?{if($_.LegacyFreeBusyStatus -eq 'OOF' -and ($_.IsAllDayEvent -eq $true -or ($_.end -  $_.start).TotalHours -gt 7)) {
                  
                $appointments = $_ ; $UserAppointments += $_; $EWSTZ += $EWSTimeZone; $vacRecord += $mailBoxes[$i]} 
                } 
            }catch {
                #$mailBoxes[$i].mail
                $recErrors += $mailBoxes[$i]
                $logFile += ",FindAppointmentsError," + $Error[0]
                $logFileErrors = $true
            }

        $NewEventFlag = $true
        $foundEvents = @()
        if ($appointments -ne $null) {
            try {
                $firstEvent = $appointments | sort-object {[datetime](Get-date($_.Start))} |select -First 1
                $lastEvent = $appointments |  sort-object {[datetime](Get-date($_.End))} -Descending |select -First 1
                $firstEventStartDelta = ($d_start - $firstEvent.start).TotalHours
                $lastEventEndDelta = ($d_end - $lastEvent.end).TotalHours
            
           #Appropriate user vacation events exist
                if(($firstEventStartDelta -ge -48 -and $lastEventEndDelta -lt 48)) {$NewEventFlag = $false; $NoNeedEventCount += 1; $logFile += ",UserEventExists"
                    $bodyTmp += '<br>' + $mailBoxes[$i].SAMAccountNameAD + "; " + $mailBoxes[$i].FullString
                    $bodyTmp += '<br>&emsp;&ensp;' + 'UserEventCount: ' + $appointments.count + '<br>&emsp;&ensp; FirstEvent: ' + ($firstEvent | select start, end,*zone, subj*, Leg*,IsAllDayEvent)
                    $bodyTmp += '<br>&emsp;&ensp;' + 'LastEvent: ' + ($lastEvent | select start, end,*zone, subj*, Leg*,IsAllDayEvent) 
                }
            } catch {
                $logFile += ",UserEventError," + $Error[0]
                $logFileErrors = $true
            }
        }
        
    
    #add vacation event existance
        try {
        if ($NewEventFlag -eq $true) {
            #Write-alError -EntryType Information -Message ("Appoi")
            $logFile += ",EventNotExist"
            $NoEventCount += 1

            $curr_ar_status = $mailboxes[$i].mail | Get-MailboxAutoReplyConfiguration |select AutoReplyState, StartTime, EndTime, OOFEventSubject, Identity,IsValid, InternalMessage
            
            #EWS Russian-timezone-when-creating-appointment' bugfix workaround
            $cusTZ1 = [TimeZoneInfo]::CreateCustomTimeZone("Time zone to workaround a bug", $EWSTimeZone.BaseUtcOffset,
                        "Time zone to workaround a bug","Time zone to workaround a bug")
            
            $Appointment = ""
            
            $Appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment($service)
		
		    $Appointment.Start= $d_start #$StartDate
		    $Appointment.End= $d_end.AddDays(1) #$EndDate
            #$Appointment.StartTimeZone = Get-TimeZone -id "UTC"
            #$Appointment.EndTimeZone = Get-TimeZone -id "UTC"
            $Appointment.StartTimeZone = [TimeZoneInfo]$custz1
            $Appointment.EndTimeZone = [TimeZoneInfo]$custz1
           
            $Subject = "Отпуск c " + $appStart + " по " + $appEnd + " (автоматическое событие)"
            $Appointment.Subject=$Subject
            $Appointment.LegacyFreeBusyStatus = 3 #"OOF"
            $Appointment.IsAllDayEvent = $true
            $Appointment.IsReminderSet = $false
      #      $Appointment.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, 0)
      #      $Appointment.Save($CalendarFolder.DisplayName, 0)
            $logFile += ",EventCreated"
            $CreatedEventCount += 1
        }
        } catch {
            $logFile += ",EventCreationError," + $Error[0]
            $logFileErrors = $true
        }

        try {
            $mailBoxes[$i].IsEventExist = $true
            $mySubject = "В почтовый календарь добавлено событие ОТПУСК"
            $body = 'Добрый день.<br><br>В Ваш почтовый календарь добавлено событие с именем "' + $Subject + '". '
            $body += "Оно поможет организаторам собраний учитывать Ваше отсутствие на рабочем месте."
            $body += '<br><br>Накануне начала отпуска рекомендуется включить в Outlook "Автоответ (Нет на работе)" и заполнить текст автоответа. '
            $body += 'Если Автоответ будет выключен, произойдёт его автоматическое включение для внутренних отправителей со стандартным текстом: "'
            $body += $oof_def_text + '".<br><br>Приятного отдыха.<br><br>Управление процессов и культуры изменений'
            $body += '<br><br> Для информации. Текущий текст внутреннего автоответа:<br>======<br>' + $curr_ar_status.InternalMessage + '<br>======'
            #$myTo = $EmailAddress
      #      Send-alMessage -body $body -to $EmailAddress -Subj $mySubject
            $bodyTmp += '<br>' + $mailBoxes[$i].SAMAccountNameAD + "; " + $Subject

        } catch {
            $logFile += ",EventCreationMessageSendingError," + $Error[0]
            $logFileErrors = $true
        }

        } # end if

    #check autoreply status
         #Write-alError -EntryType Information -Message ("before Set")
        
       if (($now_date -ge (Get-Date($mailBoxes[$i].PregLeaveFrom))) -and ($now_date -le (Get-Date($mailBoxes[$i].PregLeaveTo))) -and $mailBoxes[$i].IsOOFOn -ne $true) {
            $logFile += ",OOF_check"
            $OOF_OffCount += 1
            #Write-alError -EntryType Information -Message ("first if")
            try {
            $curr_ar_status = $mailboxes[$i].mail | Get-MailboxAutoReplyConfiguration |select AutoReplyState, StartTime, EndTime, OOFEventSubject, Identity,IsValid, InternalMessage
            
            $StartTimeDelta = ((Get-Date($mailBoxes[$i].PregLeaveFrom)).Date - (Get-Date($curr_ar_status.StartTime)).Date).Days
            $EndTimeDelta = ((Get-Date($mailBoxes[$i].PregLeaveTo)).Date - (Get-Date($curr_ar_status.EndTime)).Date).Days

            if (-not((($StartTimeDelta -ge -2 -and $EndTimeDelta -le 2 -and $curr_ar_status.AutoReplyState -eq "Scheduled"))`
             -or $curr_ar_status.AutoReplyState -eq "Enabled")) {
                #Write-alError -EntryType Information -Message ("2nd if")
                try {
                $bodyTmp += "<br>" + $mailBoxes[$i].SAMAccountNameAD + "; " + $mailBoxes[$i].mail + "; " + $mailBoxes[$i].PregLeaveFrom  + "; " + $mailBoxes[$i].PregLeaveTo + "; " + $mailBoxes[$i].PregLeaveName
                $bodyTmp += "<br>&emsp;&ensp;" + $curr_ar_status.AutoReplyState + "; " + $curr_ar_status.StartTime + "; " + $curr_ar_status.EndTime 
                #Write-alError -EntryType Information -Message ("Set")
      #          Set-MailboxAutoReplyConfiguration $mailboxes[$i].mail –InternalMessage $oof_def_text -AutoReplyState Scheduled –StartTime $d_start -EndTime $d_end.AddDays(1) -ExternalAudience None
                $mailBoxes[$i].IsOOFOn = $true
                $logFile += ",OOF_On"
                $OOF_OnCount += 1
                } catch {
                     $logFile += ",SetAutoReplyConfigurationError," + $Error[0]
                     $logFileErrors = $true
                }

                #send message to user if auto event has not been created
                if (-not($mailBoxes[$i].IsEventExist -eq $true)) {
                    try {
                    $mySubject = 'Автоответ "Нет на работе" включен (событие ОТПУСК)'
                    $body = 'Добрый день.<br><br>В Вашем почтовом ящике на период отпуска включен автоответ "Нет на работе" для внутренних отправителей со стандартным текстом: "'
                    $body += $oof_def_text + '".<br><br>Приятного отдыха.<br><br>Управление процессов и культуры изменений'
                    $body += '<br><br> Для информации. Текущий текст внутреннего автоответа:<br>======<br>' + $curr_ar_status.InternalMessage + '<br>======'
                    #$myTo = $EmailAddress
      #      Send-alMessage -body $body -to $EmailAddress -Subj $mySubject
                    $body1 += '<br>' + $mailBoxes[$i].SAMAccountNameAD + "; " + $Subject
                    $logFile += ",OOF_OnMessageSent"
                    $OOF_OnMessageSentCount += 1
                    } catch {
                        $logFile += ",OOFTurnOnMessageSendingError," + $Error[0]
                        $logFileErrors = $true
                    }
                }
            }
         } catch {
                 $logFile += ",GetAutoReplyConfigurationError," + $Error[0]
                 $logFileErrors = $true
            }   
        }
        $logFile += ";" + ($mailBoxes[$i].psobject.Properties.value -join ";") + "`r`n"
        #$mailBoxes[$i].psobject.Properties.value -join ";"

    }
    Out-File -InputObject $logFile -Encoding utf8 -FilePath $file_log
    
        
     #Write-alError -EntryType Information -Message ("Before Stop")
    
    $body1 += '<br>' + $MustEventCount + " : Count of events should be exist"
    $body1 += '<br>' + $NoEventCount + " : Count of events that are not exist"
    $body1 += '<br>' + $CreatedEventCount + " : Count of created events"
    $body1 += '<br>' + $NoNeedEventCount + " : Count of events that are not to be necessary created (appropriate user-created events have been found)"
    $body1 += '<br>' + $OOF_OffCount + " : Count of OOF-Off records should be checked"
    $body1 += '<br>' + $OOF_OnCount + " : Count of OOF turned On"
    $body1 += '<br>' + $OOF_OnMessageSentCount + " : Count of 'OOF turned On' messages have been sent to user"
    
    if($logFileErrors) {
        $body1 += '<br><br> There were some errors. Addition info logfile ' + $file_log
    } 

    
    #Write-alError -EntryType Information -Message ("Stop")

    $mailBoxes + $mailBoxes_not | select $param | Export-Csv -Path $file -Delimiter ";" -Encoding UTF8

    Write-alError -EntryType Information -Message ($body1.Replace('<br>', "`r`n"))

    $body1 += '<br>------<br>' + $bodyTmp

    Send-alMessage -body $body1 -to $myTo -Subj $mySubj
    
    Remove-alLogFiles -log_Path $logPath -logDepth 9 #remove log files older then 9 days

} catch {
      $mailBoxes + $mailBoxes_not | select $param | Export-Csv -Path $file -Delimiter ";" -Encoding UTF8
      Out-File -InputObject $logFile -Encoding utf8 -FilePath $file_log
      Write-Host $Error[0].Exception.Message
      $body =  "Error : " + $Error[0].Exception.Message
      $body += "<br>" + ($Error[0].CategoryInfo -join '`r`n')
      $body += "<br>" + $Error[0].ScriptStackTrace
      $body += "<br><br> Script execution fails. Addition info logfile " + $file_log
      Write-alError -EntryType Error -Message $body
      #Write-alError -EntryType Error -Message "Zabbix test5"
      Send-alMessage -body $body -To $myTo -Subj $mySubj
} 
