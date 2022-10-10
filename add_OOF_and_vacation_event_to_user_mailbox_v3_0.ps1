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
    Write-EventLog -LogName "Application" -Source "PSScriptOOF" -EventID 1 -EntryType $EntryType -Message $Message -Category 1 -RawData 10,20
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
<#    $connectionString = 'Server=10.46.65.206\;Database=rbstaff;Trusted_Connection=True;'
    $c_date = (Get-Date -Format 'yyyy-MM-dd')
    $query = "SELECT [SAMAccountNameAD]
                 ,[PregLeaveName]
                 ,[PregLeaveFrom]
                 ,[PregLeaveTo]
             FROM [RBStaff].[dbo].[vwShownVacation]
             where PregLeaveFrom >= '$($c_date)'"
    $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $connectionString
    $command = $connection.CreateCommand()
    $command.CommandText = $query
    $adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $command
    $dataset = New-Object -TypeName System.Data.DataSet
    $adapter.Fill($dataset)
    $memb_new_SamAccountName_list = $dataset.Tables[0]
    $memb_new_SamAccountName_list1 = $memb_new_SamAccountName_list |select SAMAccountNameAD,PregLeaveName,PregLeaveFrom,PregLeaveTo

    #$file1 = $env:USERPROFILE +"\Documents\add_OOF_psscript\20220908-190034.csv"
    $file = $env:USERPROFILE +"\Documents\add_OOF_psscript\vacation_list.csv"
    $tmp2 = Import-Csv -Path $file -Delimiter ";" -Encoding UTF8 |select SAMAccountNameAD,PregLeaveName,PregLeaveFrom,PregLeaveTo
    #$tmp2.PregLeaveTo | select -First 3
    #$tmp2 |Get-Member
    #$tmp2.Count
    $diff_email_list = Compare-Object -ReferenceObject $tmp2 -DifferenceObject $memb_new_SamAccountName_list1 # -IncludeEqual # -Property SAMAccountNameAD,PregLeaveName,PregLeaveFrom,PregLeaveTo
    #$diff_email_list |Get-Member
    #$diff_email_list |select -First 3 SideIndicator
    $vacRecAdd = $diff_email_list | ?{$_.SideIndicator -eq "<=" }
    #$vacRecOld = $diff_email_list | ?{$_.SideIndicator -ne "<=" }
    #$vacRecAdd.Count | Select -First 3
    #$vacRecOld.Count
    
    $body1 = [string]$vacRecAdd.Count + " : Count of new vacation records to add<br>------<br>Added vacation records<br>------<br>"
    $body1 += ($vacRecAdd.InputObject |Out-String).Split('`r`n')  -join '<br>'
    
    if ($vacRecAdd -ne $null) {$tmp2 += $vacRecAdd.InputObject}
    
    
 #>   
    $now_date = (Get-date).Date
    $body1 = ""
    $file = $env:USERPROFILE +"\Documents\add_OOF_psscript\vacation_list.csv"
    $file_log = $env:USERPROFILE +"\Documents\add_OOF_psscript\logs\" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt"
    #Write-alError -EntryType Information -Message ($file)

    $tmp2 = ""
    if (Test-Path -Path $file -PathType Leaf) {
        $tmp2 = Import-Csv -Path $file -Delimiter ";" -Encoding UTF8 #|select SAMAccountNameAD,PregLeaveName,PregLeaveFrom,PregLeaveTo, @{n='StartShift'; e={0}}, @{n='EndShift'; e={0}}, @{n='IsEventExist'; e={$false}}, @{n='IsOOFOn'; e={$false}}, @{n='FullString'; e={$_.SAMAccountNameAD + $_.PregLeaveName + $_.PregLeaveFrom + $_.PregLeaveTo}}
        }

    
    #$tmp2 = ""
    #[System.GC]::Collect()
    $vacRecActual = $tmp2 | ?{((Get-Date($_.PregLeaveTo) -ErrorAction SilentlyContinue).Date) -ge $now_date}
    
    #$vacRecActual = $vacRecActual | ?{($_.SAMAccountNameAD -eq 'rb075856') -or ($_.SAMAccountNameAD -eq 'rb071491') -or ($_.SAMAccountNameAD -eq 'rb101851')`
    #-or ($_.SAMAccountNameAD -eq 'rb102984') -or ($_.SAMAccountNameAD -eq 'rb076011') -or ($_.SAMAccountNameAD -eq 'rb065356') -or ($_.SAMAccountNameAD -eq 'rb073443')}

    #$vacRecActual | Export-Csv -Path $file -Delimiter ";" -Encoding UTF8

    #$vacRecActual = $vacRecActual | ?{($_.SAMAccountNameAD -eq 'rb075856')}
    
    $body1 += "<br>" + $vacRecActual.Count + " : Count of actual vacation records"
    #$body1 += ($vacRecActual |Out-String).Split('`r`n') -join '<br>'
    
    #$vacRecDel = $tmp2 | ?{((Get-Date($_.PregLeaveTo) -ErrorAction SilentlyContinue).Date) -lt $now_date.AddDays(0)}
    #Write-alError -EntryType Information -Message ("foreach next")
    #$body1 += "<br>" + $vacRecDel.Count + " : Count of vacation records to delete<br>------<br>Deleted vacation records<br>------<br>"
    #$body1 += ($vacRecDel |Out-String).Split('`r`n') -join '<br>'
    
#    $vacRecActual | Export-Csv -Path $file -Delimiter ";" -Encoding UTF8
    #Send-alMessage -body $body

    #$vacRecAdd.InputObject | Select -First 3
    #$tmp2 += $vacRecold.InputObject
    #$vacRecActual.Count

    #$vacRecOld.InputObject | Select -First 9
    
    

#get email's list
    

#get email's list at users' domain
    
    $memb_new_rosbank_email = ""
    #Write-alError -EntryType Information -Message ("select")
    $memb_new_rosbank_email = $vacRecActual | select *, @{n='mail';e={%{(Get-ADUser -Server $domain_rb -LDAPFilter "(&(SamAccountName=$($_.SAMAccountNameAD))(!userAccountControl:1.2.840.113556.1.4.803:=2))" -Properties mail -ErrorAction SilentlyContinue | select mail).mail}}}
    
    #$file = $env:USERPROFILE +"\Documents\add_OOF_psscript\vacation_list_emails.csv"
    #$memb_new_rosbank_email | Export-Csv -Path $file -Delimiter ";" -Encoding UTF8
    #($memb_new_rosbank_email.mail |Sort-Object -Unique).Count
    #($memb_new_rosbank_email | ?{$_.mail -ne $null}).Count
    
    #get existed email
    $email_exist = $memb_new_rosbank_email | ?{$_.mail -ne $null}
    #$email_exist.Count
    
    #get existed mailboxes
    #$mailbox_exist = $email_exist | select *, @{n='isMailBox';e={ if((get-mailbox $_.mail -ErrorAction SilentlyContinue) -ne $null) {1}}}
    $mailbox_exist = $memb_new_rosbank_email | select *, @{n='isMailBox';e={ if($_.mail -ne $null) {if((get-mailbox $_.mail -ErrorAction SilentlyContinue) -ne $null) {1}}}}

    #($mailbox_exist | ?{$_.isMailBox -eq 1}).count

    #$mailbox_exist |Out-GridView
    
    #get mailboxes to proceed
    $mailBoxes = $mailbox_exist | ?{$_.isMailBox -eq 1}
    $body1 += "<br>" + $mailBoxes.Count + " : Count of existed maiboxes"
    $mailBoxes_not = $mailbox_exist | ?{(-not($_.isMailBox -eq 1))}
    #$mailBoxes.Count
    #$mailBoxes_not.Count
    # ($mailBoxes + $mailBoxes_not).count

    #$mailBoxes[0].psobject.properties | select Name

    #$mailBoxes_temp = $mailBoxes | ?{$_.mail -eq 'aleksandr.ilyushenko@rosbank.ru'}
   # Get-Date($mailBox.PregLeaveFrom)
    
    #Get-date($d_start)
    #Format-alCountWorkingDays -j -5

    #$mailboxes[0]

    #$mailBoxes_ar_warning = $mailBoxes | ?{(Get-Date($_.PregLeaveFrom)) -eq (Format-alCountWorkingDays -j -5)}
    #$mailBoxes_ar_warning.Count

    #$curr_ar_status = $mailboxes.mail | Get-MailboxAutoReplyConfiguration |select AutoReplyState, StartTime, EndTime #, OOFEventSubject, Identity,IsValid, InternalMessage

    #$mailBox = $mailboxes
    #Write-alError -EntryType Information -Message ("body")
    $body = ""
    #$body1 = ""

    #$mailBoxes_temp = $mailBoxes |select -First 1
    
    $i = ""
    $logFile = "Log_flags;SAMAccountNameAD;PregLeaveName;PregLeaveFrom;PregLeaveTo;StartShift;EndShift;IsEventExist;IsOOFOn;FullString;mail;isMailBox`r`n"

    $MustEventCount = 0 #Count of events to be exist
    $CreatedEventCount = 0 #Count of events to be created
    $NoEventCount = 0 #Count of events that are not exist
    $OOF_OffCount = 0 #Count of OOF-Off records to be proceeded
    $OOF_OnCount = 0 #Count of OOF turned On

    for ($i = 0; $i -lt $mailBoxes.Count; $i++) {
         #Write-alError -EntryType Information -Message ($mailBox)
        #$body1 +='<br>' +$mailBoxes[$i].mail
        $d_start = Get-Date($mailBoxes[$i].PregLeaveFrom) # -Format 'dd.MM.yyyy'

        if (($d_start.AddDays(5) -ge $now_date) -and ($mailBoxes[$i].IsEventExist -ne $true)) {

        $logFile += "5,EventNotCreated"
        $MustEventCount += 1

        $d_end = Get-Date($mailBoxes[$i].PregLeaveTo) # -Format 'dd.MM.yyyy'
        $appStart = $d_start.ToString('dd.MM')
        $appEnd = $d_end.ToString('dd.MM')
        $oof_def_text = "Уважаемые коллеги, добрый день. С "  + $appStart + " по " + $appEnd + "  нахожусь в отпуске."

    #check vacation event existance
        
        $appointments = $null
        #$EmailAddress = "aleksandr.ilyushenko@rosbank.ru"
        ##EmailAddress = "Aleksey.A.Semenov@rosbank.ru"
        $EmailAddress = $mailBoxes[$i].mail
        #$EWSURL = '' #gets from init parameters' file
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
        $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
        $Service.Url = [System.URI] $EWSURL
        $service.UseDefaultCredentials = $true
        #$service.HttpHeaders.Add("X-AnchorMailbox", $EmailAddress)

        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress)
        $CalendarFolder = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)

        $View = [Microsoft.Exchange.WebServices.Data.CalendarView]::new($d_start,$d_end.AddDays(1))
        $View.PropertySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject,
                             [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start,[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End,
                             [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::DateTimeSent)
        #$View.PropertySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::new()
        $appointments = $CalendarFolder.FindAppointments($View) | ?{$_.Subject -like 'Отпуск*(автоматическое событие)'}# |  Select-Object start , End , Subject 
        #Write-alError -EntryType Information -Message ("17")
        
    
    #add vacation event existance
        if ($appointments -eq $null) {
            #Write-alError -EntryType Information -Message ("Appoi")
            $logFile += ",EventNotExist"
            $NoEventCount += 1

            $curr_ar_status = $mailboxes[$i].mail | Get-MailboxAutoReplyConfiguration |select AutoReplyState, StartTime, EndTime, OOFEventSubject, Identity,IsValid, InternalMessage
            
            $Appointment = ""
            
            $Appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment($service)
		
		    $Appointment.Start= $d_start #$StartDate
		    $Appointment.End= $d_end.AddDays(1) #$EndDate
            $Appointment.StartTimeZone = Get-TimeZone -id "UTC"
            $Appointment.EndTimeZone = Get-TimeZone -id "UTC"
            #$Appointment.StartTimeZone = $MoscowTimeZone
            #$Appointment.EndTimeZone = $MoscowTimeZone
            #$appStart = $Appointment.Start.ToString('dd.MM')
            #$appEnd = $Appointment.End.AddDays(-1).ToString('dd.MM')

            #$oof_def_text = "Уважаемые коллеги, добрый день. С "  + $appStart + " по " + $appEnd + "  нахожусь в отпуске."
            $Subject = "Отпуск c " + $appStart + " по " + $appEnd + " (автоматическое событие)"
            $Appointment.Subject=$Subject
            $Appointment.LegacyFreeBusyStatus = 3 #"OOF"
            $Appointment.IsAllDayEvent = $true
            $Appointment.IsReminderSet = $false
      #      $Appointment.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, 0)
            $logFile += ",EventCreated"
            $CreatedEventCount += 1

            $mailBoxes[$i].IsEventExist = $true
            $mySubject = "В почтовый календарь добавлено событие ОТПУСК"
            $body = 'Добрый день.<br><br>В Ваш почтовый календарь добавлено событие с именем "' + $Subject + '". '
            $body += "Оно поможет организаторам собраний учитывать Ваше отсутствие на рабочем месте."
            $body += '<br><br>Накануне начала отпуска рекомендуется включить в Outlook "Автоответ (Нет на работе)" и заполнить текст автоответа. '
            $body += 'Если Автоответ будет выключен, произойдёт его автоматическое включение для внутренних отправителей со стандартным текстом: "'
            $body += $oof_def_text + '".<br><br>Приятного отдыха.'
            $body += '<br><br> Для информации. Текущий текст внутреннего автоответа:<br>======<br>' + $curr_ar_status.InternalMessage + '<br>======'
            #$myTo = $EmailAddress
      #      Send-alMessage -body $body -to $EmailAddress -Subj $mySubject
            $body1 += '<br>' + $Subject

        }
        }
        
        #$Appointments.Start.ToString('dd.MM')
        #$Appointments.End.ToString('dd.MM')
        #$Appointments.Subject
        #$curr_ar_status = ""
    #check autoreply status
         #Write-alError -EntryType Information -Message ("before Set")
        
        #$body+="<br>" + (($now_date -ge (Get-Date($mailBox.PregLeaveFrom))) -and ($now_date -le (Get-Date($mailBox.PregLeaveTo)))).ToString()


        if (($now_date -ge (Get-Date($mailBoxes[$i].PregLeaveFrom))) -and ($now_date -le (Get-Date($mailBoxes[$i].PregLeaveTo))) -and $mailBoxes[$i].IsOOFOn -ne $true) {
            $logFile += ",OOF_Off"
            $OOF_OffCount += 1
            #Write-alError -EntryType Information -Message ("first if")
            $curr_ar_status = $mailboxes[$i].mail | Get-MailboxAutoReplyConfiguration |select AutoReplyState, StartTime, EndTime, OOFEventSubject, Identity,IsValid, InternalMessage
            $StartTimeDelta = [Math]::Abs(((Get-Date($mailBoxes[$i].PregLeaveFrom)).Date - (Get-Date($curr_ar_status.StartTime)).Date).Days)
            $EndTimeDelta = [Math]::Abs(((Get-Date($mailBoxes[$i].PregLeaveTo)).Date - (Get-Date($curr_ar_status.EndTime)).Date).Days)
            if (-not((($StartTimeDelta -le 2 -and $EndTimeDelta -le 2 -and $curr_ar_status.AutoReplyState -eq "Scheduled"))`
             -or $curr_ar_status.AutoReplyState -eq "Enabled")) {
                #Write-alError -EntryType Information -Message ("2nd if")
                $body1 += "<br>" + $mailBoxes[$i].SAMAccountNameAD + "; " + $mailBoxes[$i].mail + "; " + $mailBoxes[$i].PregLeaveFrom  + "; " + $mailBoxes[$i].PregLeaveTo + "; " + $mailBoxes[$i].PregLeaveName
                $body1 += "<br>&emsp;&ensp;" + $curr_ar_status.AutoReplyState + "; " + $curr_ar_status.StartTime + "; " + $curr_ar_status.EndTime 
                #Write-alError -EntryType Information -Message ("Set")
      #          Set-MailboxAutoReplyConfiguration $mailboxes[$i].mail –InternalMessage $oof_def_text -AutoReplyState Scheduled –StartTime $d_start -EndTime $d_end.AddDays(1) -ExternalAudience None
                $mailBoxes[$i].IsOOFOn = $true
                $logFile += ",OOF_On"
                $OOF_OnCount += 1
            }
            
        }
        $logFile += ";" + ($mailBoxes[$i].psobject.Properties.value -join ";") + "`r`n"
        #$mailBoxes[$i].psobject.Properties.value -join ";"

    }
    Out-File -InputObject $logFile -Encoding utf8 -FilePath $file_log
    
        
<#
        #$body = ""
        if (((Get-Date($mailBox.PregLeaveFrom)) -eq (Format-alCountWorkingDays -j -5))) {
            #$mySubject = "Установить автоовет об ОТПУСК"
            #$body = 'Добрый день.<br><br> "' + $Subject + '". '
            #$body += "Оно поможет организаторам собраний учитывать Ваше осутствие на рабочем месте.<br><br>Приятного отдыха."
            $body += "<br>" + $mailBox.SAMAccountNameAD + "; " + $mailBox.mail + "; " + $mailBox.PregLeaveFrom  + "; " + $mailBox.PregLeaveTo + "; " + $mailBox.PregLeaveName
            #Send-alMessage -body $body
        }
        
#>        
     #Write-alError -EntryType Information -Message ("Before Stop")
    
    $body1 += '<br>' + $MustEventCount + " : Count of events to be exist"
    $body1 += '<br>' + $CreatedEventCount + " : Count of events to be created"
    $body1 += '<br>' + $NoEventCount + " : Count of events that are not exist"
    $body1 += '<br>' + $OOF_OffCount + " : Count of OOF-Off records to be proceeded"
    $body1 += '<br>' + $OOF_OnCount + " : Count of OOF turned On"


    Send-alMessage -body $body1 -to $myTo -Subj $mySubj
    #Write-alError -EntryType Information -Message ($body1 + "`r`n" + $myTo + "`r`n" + $mySubj)
    #Write-alError -EntryType Information -Message ("Stop")
    $mailBoxes + $mailBoxes_not | select SAMAccountNameAD,PregLeaveName,PregLeaveFrom,PregLeaveTo, StartShift, EndShift, IsEventExist, IsOOFOn, FullString | Export-Csv -Path $file -Delimiter ";" -Encoding UTF8

} catch {
      $mailBoxes + $mailBoxes_not | select SAMAccountNameAD,PregLeaveName,PregLeaveFrom,PregLeaveTo, StartShift, EndShift, IsEventExist, IsOOFOn, FullString | Export-Csv -Path $file -Delimiter ";" -Encoding UTF8
      Out-File -InputObject $logFile -Encoding utf8 -FilePath $file_log
      Write-Host $Error[0].Exception.Message
      $body =  "Error : " + $Error[0].Exception.Message
      $body += "<br>" + ($Error[0].CategoryInfo -join '`r`n')
      $body += "<br>" + $Error[0].ScriptStackTrace
      $body += "<br><br> Script execution fails. No changes were made."
      Write-alError -EntryType Error -Message $body
      Send-alMessage -body $body -To $myTo -Subj $mySubj
} 
