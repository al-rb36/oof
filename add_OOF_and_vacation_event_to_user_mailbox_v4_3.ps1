#This script
# a) creates VACATION mailbox event with Out-of-Office time mark
# b) turn OOF on if necessary.
#
# Verison 4.3
# (c) Aleksandr Ilyushenko, 2023
#

#Exit


[Reflection.Assembly]::LoadFile("C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll") | Out-Null
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

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

Function Get-alEWSEvent {

        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1

        $TZ_default = [System.TimeZoneInfo](Get-TimeZone)
        
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
                             [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::TimeZone,
                             [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Organizer)
          #                   [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::StartTimeZone,
          #                   [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::EndTimeZone)
        
        #$appointments = $CalendarFolder.FindAppointments($View) | ?{$_.Subject -like 'Отпуск*(автоматическое событие)'}# |  Select-Object start , End , Subject 
        #$appointments = $CalendarFolder.FindAppointments($View) | ?{($_.LegacyFreeBusyStatus -eq 'OOF') -and ($_.IsAllDayEvent -eq $true) }# |  Select-Object start , End , Subject 
        #Write-alError -EntryType Information -Message ("17")
            try {  
                $appointments = $CalendarFolder.FindAppointments($View)  
            }catch {
                #$mailBoxes[$i].mail
                $recErrors += $mailBoxes[$i]
                $logFile += ",FindAppointmentsError," + $Error[0]
                $logFileErrors = $true
            }
            return [PSCustomObject]@{ 
                'EWSTimeZone' = $EWSTimeZone;
                'service' = $service;
                'appointments' = $appointments;
                }

}

try {

#load additional functions
    ."C:\Scripts\al_PSScripts\alPSScript_library_v_1.ps1"

#load init parameters' file
    ."C:\Scripts\al_PSScripts\add_OOF_and_vacation_event_to_user_mailbox\OOF_init.ps1"

#load default message to user 
    ."C:\Scripts\al_PSScripts\add_OOF_and_vacation_event_to_user_mailbox\oof_def_text.ps1"

#Init data. If it is not defined, gets from init parameters' file
    $PS_Script_Tilte = ""
    $log_msg = @{}
    $c_in =""
    #$domain_rb = ""
    #$domain_gts = ""
    #$domain_gts_SearchBase = ""
    #$connectionString = ""
    #$EWSURL = '' #
    $logPath = $env:USERPROFILE +"\Documents\add_OOF_psscript\logs\"

# test ONLY
<#         $tz1 = Get-TimeZone -id (Get-MailboxRegionalConfiguration $myTo).TimeZone
            $EWSTimeZone_al = [System.TimeZoneInfo]::FindSystemTimeZoneById($tz1.id)
        } catch {
            $EWSTimeZone_al = $TZ_default
        }
$Service_al = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion,$EWSTimeZone_al)
$Service_al.Url = [System.URI] $EWSURL
$service_al.UseDefaultCredentials = $true
$service_al.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $myTo)    
$service_al.UseDefaultCredentials = $true
#>   
    #Send-myMessage constants
    #$myFrom = ""
    #$mySubj = ""
    #$mySmtpServer = ""
    #$myTo = ""

    #check event log config
    #$ApplicationName = "MLS PSScripts"
    #$EvenSourceName = "PSScriptOOF_MLS"
    New-alEventLog -ApplicationName $ApplicationName -EvenSourceName $EvenSourceName

#get current vacation data
  
    $now_date = (Get-date).Date
    $body1 = ""
    $bodyTmp = ""
    #$file = $env:USERPROFILE +"\Documents\add_OOF_psscript\20221129 _vacation-event_test_list.csv"
    #actual vacations data list from the Source
    $file = "C:\Users\rbs_MailScript" +"\Documents\add_OOF_psscript\vacation_list_$odbcSourceName.csv"
    #extedned vacation info
    $file_ext = $env:USERPROFILE + "\Documents\add_OOF_psscript\vacation_list_$($odbcSourceName)_ext.csv"
    #Organizer info service
    $file_ois = $env:USERPROFILE + "\Documents\add_OOF_psscript\vacation_list_$($odbcSourceName)_ois.csv"
    #execute the script for selected users only
    $fileUserLimiter = $env:USERPROFILE + "\Documents\add_OOF_psscript\vacation_list_$($odbcSourceName)_userlimiter.csv"

    $file_log = $logPath + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt"
    #Write-alError -EntryType Information -Message ($file)

    $param_ext = New-Object System.Collections.Specialized.OrderedDictionary
    $param_ext.Add('bss_employee_vacation_id', '')
    $param_ext.Add('IsEventExist', '')
    $param_ext.Add('IsOOFOn', '')
    $param_ext.Add('FullString', '')
    $param_ext.Add('employee_login', '')
    $param_ext.Add('dt_vacation_from', '')
    $param_ext.Add('dt_vacation_to', '')
    $param_ext.Add('is_deleted', '')
    $param_ext.Add('dt_last_update', '')
    $param_ext.Add('EventID', '')

 <#   
    $param_ext = @{bss_employee_vacation_id = ''
               IsEventExist = ''
               IsOOFOn =''
               FullString = ''
               employee_login = ''
               dt_vacation_from = ''
               dt_vacation_to = ''
               is_deleted = ''
               dt_last_update = ''
               EventID = ''}
#>

    $param_ext.dt_vacation_to = $now_date.ToString()
    
   [PSObject[]]$vacRecActualExt = New-Object -TypeName PSObject -Property $param_ext

#read extedned vacation info
    #$param_ext = '"bss_employee_vacation_id";"IsEventExist";"IsOOFOn";"FullString";"employee_login";"dt_vacation_from";"dt_vacation_to";"is_deleted";"dt_last_update"'
    #read vacation ext info file if it is exist
    if (Test-Path -Path $file_ext -PathType Leaf) {
        $vacRecActualExt = Import-Csv -Path $file_ext -Delimiter ";" -Encoding UTF8 
    }

    #$vacRecActualExt | Out-GridView


    $tmp2 = ""
    $param = @()
    if (Test-Path -Path $file -PathType Leaf) {
        $tmp2 = Import-Csv -Path $file -Delimiter ";" -Encoding UTF8 #|select SAMAccountNameAD,PregLeaveName,PregLeaveFrom,PregLeaveTo, @{n='StartShift'; e={0}}, @{n='EndShift'; e={0}}, @{n='IsEventExist'; e={$false}}, @{n='IsOOFOn'; e={$false}}, @{n='FullString'; e={$_.SAMAccountNameAD + $_.PregLeaveName + $_.PregLeaveFrom + $_.PregLeaveTo}}
        $param = $tmp2[0].psobject.Properties.Name
        }

                          # for PROD purposes only
    $vacRecActual = @()
    $vacRecActual = $tmp2 | ?{((Get-Date($_.dt_vacation_to) -ErrorAction SilentlyContinue).Date) -ge $now_date}
    
    $tmp2[1].Count
    $vacRecActual.Count

#use script for selected users only (prod test script execution)
     $userLimiter = @()
     #$body1 += "<br>" + (Test-Path -Path $fileUserLimiter -PathType Leaf)
     if (Test-Path -Path $fileUserLimiter -PathType Leaf) {
        $userLimiter = Import-Csv -Path $fileUserLimiter -Delimiter ";" -Encoding UTF8
        
        }
    $test01 = @()
    $vacRecActual | %{ if($userLimiter.employee_login -eq $_.employee_login) {$test01 += $_ }}
    
    #$test01 |Out-GridView
    $vacRecActual = $test01

    $body1 += "<br>" + $vacRecActual.Count + " : Count of actual vacation records"
   
#get email's list
    

#get email's list at users' domain
    
    [PSObject[]]$memb_new_rosbank_email = @()
    [PSObject[]]$email_exist = @()
    [PSObject[]]$mailbox_exist = @()
    #Write-alError -EntryType Information -Message ("select")
    $memb_new_rosbank_email = $vacRecActual | select *, @{n='mail';e={%{(Get-ADUser -Server $domain_rb -LDAPFilter "(&(SamAccountName=$($_.employee_login))(!userAccountControl:1.2.840.113556.1.4.803:=2))" -Properties mail -ErrorAction SilentlyContinue | select mail).mail}}}
    #$body1 += "<br>" + ($memb_new_rosbank_email.count)
    #get existed email
    $email_exist = $memb_new_rosbank_email | ?{$_.mail -ne $null}
    #$email_exist.Count
    #$body1 += "<br>" + ($email_exist.count)
    #mark if mailbox exists
    $mailbox_exist = $memb_new_rosbank_email | select *, @{n='isMailBox';e={ if($_.mail -ne $null) {if((get-mailbox $_.mail -ErrorAction SilentlyContinue) -ne $null) {1}}}}

    $mailbox_exist.Count

    #$mailbox_exist | Out-GridView

    #get mailboxes to proceed
    [PSObject[]]$mailBoxes = @()
    $mailBoxes = $mailbox_exist | ?{$_.isMailBox -eq 1}
    $body1 += "<br>" + $mailBoxes.Count + " : Count of existed maiboxes"
    $mailBoxes_not = $mailbox_exist | ?{(-not($_.isMailBox -eq 1))}
    $mailBoxes.Count
    #$mailBoxes_not.Count

    #$mailBoxes | Out-GridView

    #$mailBoxes_not |Out-GridView

    #Write-alError -EntryType Information -Message ("body")
    $body = ""
    #$body1 = ""

    #$mailBoxes_temp = $mailBoxes |select -First 1

    #synchronize actual vacation data with extended data of vacation info

<#
    for ($i = 0; $i -lt $mailBoxes.Count; $i++) {
        #$mailBoxes[$i].bss_employee_vacation_id | % {if (-not($_ -eq $vacRecActualExt.bss_employee_vacation_id)) { $j = ($_) }}
        if ([array]::indexof($vacRecActualExt.bss_employee_vacation_id,$mailBoxes[$i].bss_employee_vacation_id) -eq -1) {
            $vacRecActualExt += [PSCustomObject]@{bss_employee_vacation_id = $mailBoxes[$i].bss_employee_vacation_id
               IsEventExist = ''
               IsOOFOn =''
               FullString = $mailBoxes[$i].bss_employee_vacation_id + $mailBoxes[$i].employee_login + $mailBoxes[$i].dt_vacation_from + $mailBoxes[$i].dt_vacation_to + $mailBoxes[$i].employee_vacation_type_nm
               employee_login = $mailBoxes[$i].employee_login
               dt_vacation_from = $mailBoxes[$i].dt_vacation_from
               dt_vacation_to = $mailBoxes[$i].dt_vacation_to
               is_deleted = $mailBoxes[$i].is_deleted
               dt_last_update = $mailBoxes[$i].dt_last_update
               EventID = ''}
        }
    }
#>
    for ($i = 0; $i -lt $mailBoxes.Count; $i++) {
        if ([array]::indexof($vacRecActualExt.bss_employee_vacation_id,$mailBoxes[$i].bss_employee_vacation_id) -eq -1) {
               $param_ext.bss_employee_vacation_id = $mailBoxes[$i].bss_employee_vacation_id
               $param_ext.FullString = $mailBoxes[$i].bss_employee_vacation_id + $mailBoxes[$i].employee_login + $mailBoxes[$i].dt_vacation_from + $mailBoxes[$i].dt_vacation_to + $mailBoxes[$i].employee_vacation_type_nm
               $param_ext.employee_login = $mailBoxes[$i].employee_login
               $param_ext.dt_vacation_from = $mailBoxes[$i].dt_vacation_from
               $param_ext.dt_vacation_to = $mailBoxes[$i].dt_vacation_to
               $param_ext.dt_last_update = $mailBoxes[$i].dt_last_update
               [array]$vacRecActualExt += New-Object -TypeName PSObject -Property $param_ext
        }
    }

  #  $vacRecActualExt | Export-Csv -Path $file_ext -Delimiter ";" -Encoding UTF8 -NoTypeInformation

# eliminate events which do not have to be controlled    
     [PSObject[]]$tmp3 = @()
     for ($i = 0; $i -lt $vacRecActualExt.Count; $i++) {
        if (-not($vacRecActualExt[$i].IsEventExist -eq 2)) {
               $j = ""
               $j = [array]::indexof($mailBoxes.bss_employee_vacation_id,$vacRecActualExt[$i].bss_employee_vacation_id)
               if ($j -ge 0) {$tmp3 += $mailBoxes[$j]}
        }
    }
    $mailBoxes = $tmp3
    #$body1 += "<br>" + ($vacRecActualExt.Count)
    #$tmp3 |Out-GridView
    #$vacRecActualExt | Out-GridView
    
    $logFile = "Log_flags;" + ($param -join ";") + ";mail;isMailBox`r`n"
    $logFileErrors = $false

    $MustEventCount = 0 #Count of events should be exist
    $CreatedEventCount = 0 #Count of events to be created
    $DeletedEventCount = 0 #Count of events to be deleted
    $UpdatedEventCount = 0 #Count of events to be updated
    $NoEventCount = 0 #Count of events that are not exist
    $NoNeedEventCount = 0 #Count of events that are not to be necessary created
    $DeletedByUserEventCount = 0 #Count of events has been deleted by user
    $OOF_OffCount = 0 #Count of OOF-Off records should be checked
    $OOF_OnCount = 0 #Count of OOF turned On
    $OOF_OnMessageSentCount = 0 #Count of 'OOF turned On' messages have been sent to user 
    $OOF_OnControlOff = 0 # #Count of OOF that do not need to be activated
    $OrganizerMessageSent = 0 #Count of messages has been sent to meeting organizers

    $TZ_default = [System.TimeZoneInfo](Get-TimeZone)
    $userAppointments =@()
    $EWSTZ = @()
    $vacRecord = @()
    $recErrors = @()

    $organizer_list = @()

#test ONLY
   # $userAppointments_al =@()
   # $EWSTZ_al = @()
   # $vacRecord_al = @()
   # $recErrors_al = @()

    #$mailBoxes =  $mailBoxes | ?{$_.IsOOFOn -eq $true -and $_.mail -ne 'Tatyana.Kulieva@rosbank.ru'}
    #$mailBoxes =  $mailBoxes | ?{$now_date -le (Get-Date($_.PregLeaveTo))}
    

    #(0..($mailboxes.Count-1)) |  % {if ($mailboxes[$_].employee_login -eq 'rb082771') {$_}}
    #(0..($mailboxes.Count-1)) |  %{$mailboxes[$_].IsOOFOn = $false


#$i = 27 2307 20 4
    for ($i = 0; $i -lt $mailBoxes.Count; $i++) {
         #Write-alError -EntryType Information -Message ($mailBox)

        $j = ""
        $j = [array]::indexof($vacRecActualExt.bss_employee_vacation_id,$mailBoxes[$i].bss_employee_vacation_id)
        
        $d_start = Get-Date($mailBoxes[$i].dt_vacation_from) # -Format 'dd.MM.yyyy'
        $d_end = Get-Date($mailBoxes[$i].dt_vacation_to) # -Format 'dd.MM.yyyy'

        $appointments = $null
        
        #$MustEventCount += 1
        #$mailBoxes[$i].IsEventExist = $false
        #$mailBoxes[$i].IsOOFOn = $false

        $appStart = $d_start.ToString('dd.MM')
        $appEnd = $d_end.ToString('dd.MM')
 
        $Subject = "Отпуск c $appStart по $appEnd (" + $mailBoxes[$i].bss_employee_vacation_id +")"

        $tt1 = (Get-Date($d_end.AddDays(1)) -Format "dd.MM.yyyy")
        $oof_def_text = "{0} $appStart {1} $appEnd {2} $tt1 {3}" -f $oof_def_text_values

        $EmailAddress = $mailBoxes[$i].mail
        
        #Exception calling "FindAppointments" with "1" argument(s): "The specified view range exceeds the maximum range of two years."
        if (($d_end - $d_start).Days -gt 600) { $vacRecActualExt[$j].IsEventExist = 1;  $MustEventCount -= 1 } # elimination of  '2 years over' error 

        #AddDays sets count of the days before event  creation
        $nowTimeDelta = ($d_start.AddDays(-90) - $now_date).Days
        if ($nowTimeDelta -le 0) { $MustEventCount += 1; $logFile += "EventMustExist" }
        
        #$body1 += "<br>" + $nowTimeDelta
        
        #vacation event existance control
        #if (($nowTimeDelta -le 0) -and ($vacRecActualExt[$j].IsEventExist -ne $true) -and (((Get-Date($mailBoxes[$i].dt_vacation_to) -ErrorAction SilentlyContinue).Date) -ge $now_date)) {
        if (($nowTimeDelta -le 0) -and ($vacRecActualExt[$j].IsEventExist -ne 1) -and (((Get-Date($mailBoxes[$i].dt_vacation_to) -ErrorAction SilentlyContinue).Date) -ge $now_date)) {

        $logFile += ",5,EventNotCreated"
     <#   
        $appStart = $d_start.ToString('dd.MM')
        $appEnd = $d_end.ToString('dd.MM')
        #$oof_def_text = "Уважаемые коллеги, добрый день. С "  + $appStart + " по " + $appEnd + "  нахожусь в отпуске."
        $oof_def_text = '«Уважаемый коллега, добрый день. С '  + $appStart + ' по ' + $appEnd + ' я нахожусь в отпуске '
        $oof_def_text += 'с ограниченным доступом к электронной почте. На ваше сообщение смогу ответить начиная с ' + (Get-Date($d_end.AddDays(1)) -Format "dd.MM.yyyy")
        $oof_def_text += '<br>Спасибо за ваше письмо»'
     #>
        #to check vacation event existance
        
        $res1 = ""
        $res1 = Get-alEWSEvent
        $appointments = $res1.appointments | ?{($_.LegacyFreeBusyStatus -eq 'OOF'  -and ($_.IsAllDayEvent -eq $true -or ($_.end -  $_.start).TotalHours -gt 7))}
        $EWSTimeZone = $res1.EWSTimeZone
        $service = $res1.service

        #$res1.appointments |Out-GridView
       
        #$EmailAddress = $mailBoxes[$i].mail

<#
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
#>

        #$NewEventFlag = $true
        if ($vacRecActualExt[$j].IsEventExist -ne 2) {$vacRecActualExt[$j].IsEventExist = 0}
        $foundEvents = @()
        if ($appointments -ne $null) {
            try {
                $firstEvent = $appointments | sort-object {[datetime](Get-date($_.Start))} |select -First 1
                $lastEvent = $appointments |  sort-object {[datetime](Get-date($_.End))} -Descending |select -First 1
                $firstEventStartDelta = ($d_start - $firstEvent.start).TotalHours
                $lastEventEndDelta = ($d_end - $lastEvent.end).TotalHours
            
           #Appropriate user vacation events exist
                #if(($firstEventStartDelta -ge -48 -and $lastEventEndDelta -lt 48)) {$NewEventFlag = $false; $NoNeedEventCount += 1; $logFile += ",UserEventExists"
                if(($firstEventStartDelta -ge -48 -and $lastEventEndDelta -lt 48)) {
                    #$NewEventFlag = $false; $vacRecActualExt[$j].IsEventExist = 2; $NoNeedEventCount += 1; $logFile += ",UserEventExists"
                    $vacRecActualExt[$j].IsEventExist = 2; $NoNeedEventCount += 1; $logFile += ",UserEventExists"
                    $bodyTmp += '<br>' + $mailBoxes[$i].employee_login + "; " + $vacRecActualExt[$j].FullString
                    $bodyTmp += '<br>&emsp;&ensp;' + 'UserEventCount: ' + $appointments.count + '<br>&emsp;&ensp; FirstEvent: ' + ($firstEvent | select start, end,*zone, subj*, Leg*,IsAllDayEvent)
                    $bodyTmp += '<br>&emsp;&ensp;' + 'LastEvent: ' + ($lastEvent | select start, end,*zone, subj*, Leg*,IsAllDayEvent) 
                }
            } catch {
                $logFile += ",UserEventSearchError," + $Error[0]
                $logFileErrors = $true
            }
        }
        
    #delete vacation event
        if (($mailboxes[$i].is_deleted -eq 1) -and ($vacRecActualExt[$j].IsEventExist -eq 1) -and (-not($vacRecActualExt[$j].is_deleted -eq 1))) {
            try {
                $App = ""
                #$res_al = ""
                #$res_al = Get-alEWSEvent_test
                #$appointments_al = $res_al.Appointments_al | ?{($_.LegacyFreeBusyStatus -eq 'OOF'  -and ($_.IsAllDayEvent -eq $true -or ($_.end -  $_.start).TotalHours -gt 7))}
                #$EWSTimeZone_al = $res_al.EWSTimeZone_al
                #$service_al = $res_al.service_al

                 #$appointments = $appointments |? {$_.Id.ToScting() -eq $vacRecActualExt[$j].EventID}
                #$appointments.Id
                #$app_i_str = ""
#test ONLY
                #$app_i_str = $appointments.Id.ToString()
                #$app_i_str = ($appointments_al |? {$_.Subject -eq $Subject}).Id.ToString()
                #$vacRecActualExt[$j].EventID = $app_i_str
                #$App = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service_al, $vacRecActualExt[$j].EventID)
                
                $App = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service, $vacRecActualExt[$j].EventID)
                $app.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems,`
                [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone)

                $logFile += ",EventDeleted"
                $DeletedEventCount += 1
                $vacRecActualExt[$j].is_deleted = 1
                $vacRecActualExt[$j].dt_last_update = $now_date
                $bodyTmp += '<br>&emsp;&ensp;' + 'EventDeleted' 
            } catch {
                $logFile += ",EventDetetionError," + $Error[0]
                $logFileErrors = $true
            }
        }

     #update vacation event
       # if ($vacRecActualExt[$j].dt_last_update -eq '') {$vacRecActualExt[$j].dt_last_update = $now_date.AddDays(-10)}
       
            try {
                 if (($mailboxes[$i].is_deleted -ne 1)`
                     -and ([datetime](Get-date($mailboxes[$i].dt_last_update)) -gt [datetime](Get-date($vacRecActualExt[$j].dt_last_update)))`
                     -and ($vacRecActualExt[$j].IsEventExist -eq 1)`
                    -and (-not($vacRecActualExt[$j].is_deleted -eq 1))) {
                    
                $App = ""
                #$res_al = ""
                #$res_al = Get-alEWSEvent_test
                #$appointments_al = $res_al.Appointments_al | ?{($_.LegacyFreeBusyStatus -eq 'OOF'  -and ($_.IsAllDayEvent -eq $true -or ($_.end -  $_.start).TotalHours -gt 7))}
                #$EWSTimeZone_al = $res_al.EWSTimeZone_al
                #$service_al = $res_al.service_al
                
                $App = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service, $vacRecActualExt[$j].EventID)

                $App.Start = $d_start
                $App.End = $d_end
                $Subject = "Отпуск c " + $appStart + " по " + $appEnd
                $App.Subject = $Subject #+ " " + $mailBoxes[$i].employee_login

                $app.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite,`
                [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::SendToNone)

                $logFile += ",EventUpdated"
                $UpdatedEventCount += 1
                $vacRecActualExt[$j].dt_last_update = $now_date
                $bodyTmp += '<br>&emsp;&ensp;' + 'EventUpdated : ' + $App.Start + ' to ' + $mailboxes[$i].dt_vacation_from + ' and ' + $App.Start + ' to ' + $mailboxes[$i].dt_vacation_to
                
                $vacRecActualExt[$j].dt_vacation_from = $mailboxes[$i].dt_vacation_from
                $vacRecActualExt[$j].dt_vacation_to = $mailboxes[$i].dt_vacation_to
                $vacRecActualExt[$j].dt_last_update = $mailboxes[$i].dt_last_update
                }
            } catch {
                $logFile += ",EvenUpdatingError," + $Error[0]
                $logFileErrors = $true
            }
        
    
    
    #to add vacation event
        
        if (($vacRecActualExt[$j].IsEventExist -eq 0) -and (-not($mailboxes[$i].is_deleted -eq 1)) ) {
         try {
        
                #Write-alError -EntryType Information -Message ("Appoi")
                $logFile += ",EventNotExist"
                $NoEventCount += 1

                $curr_ar_status = $mailboxes[$i].mail | Get-MailboxAutoReplyConfiguration |select AutoReplyState, StartTime, EndTime, OOFEventSubject, Identity,IsValid, InternalMessage
            
                #'EWS Russian-timezone-when-creating-appointment' bugfix workaround
                $cusTZ1 = [TimeZoneInfo]::CreateCustomTimeZone("Time zone to workaround a bug", $EWSTimeZone.BaseUtcOffset,
                          "Time zone to workaround a bug","Time zone to workaround a bug")
            
                $App = ""
                $Appointment = ""
                $Appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment($service)
                #$Appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment($service_al)
		        $Appointment.Start = $d_start #$StartDate
    		    $Appointment.End = $d_end.AddDays(1) #$EndDate
                #$Appointment.StartTimeZone = Get-TimeZone -id "UTC"
                #$Appointment.EndTimeZone = Get-TimeZone -id "UTC"
                $Appointment.StartTimeZone = [TimeZoneInfo]$custz1
                $Appointment.EndTimeZone = [TimeZoneInfo]$custz1
                #$Subject = "Отпуск c " + $appStart + " по " + $appEnd + " (автоматическое событие)"
                #$Subject = " c " + $appStart + " по " + $appEnd
                $Appointment.Subject = $Subject
                $Appointment.LegacyFreeBusyStatus = 3 #"OOF"
                $Appointment.IsAllDayEvent = $true
                $Appointment.IsReminderSet = $false
                $Appointment.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, 0)
                #$Appointment.Save($CalendarFolder.DisplayName, 0)
                
                $res1 = ""
                $res1 = Get-alEWSEvent
                $appointments = $res1.Appointments | ?{($_.LegacyFreeBusyStatus -eq 'OOF'  -and ($_.IsAllDayEvent -eq $true -or ($_.end -  $_.start).TotalHours -gt 7))}
                $EWSTimeZone = $res1.EWSTimeZone
                $service = $res1.service
                $res1.appointments |Out-GridView

                #$res_al.appointments_al <#| select Start, End, Organizer, Subject #>|Out-GridView

                #$res_al.appointments_al.Organizer | Sort-Object -Unique |Out-GridView

#Get-Mailbox $organizer_list.organizer_address
                #$res_al.appointments_al.TotalCount
                #$appointments_al.Count

                # | ?{if($_.LegacyFreeBusyStatus -eq 'OOF'  -and ($_.IsAllDayEvent -eq $true -or ($_.end -  $_.start).TotalHours -gt 7)) {
               # 
               # $appointments_al = $_ ; $UserAppointments_al += $_; $EWSTZ_al += $EWSTimeZone_al; $vacRecord_al += $mailBoxes[$i]} 
               # } 
                
                #$appointments = $appointments_al |? {$_.Subject -eq $Subject}
                #$appointments.Id
                $app_i_str = ""

                $app_i_str = ($appointments |? {$_.Subject -eq $Subject}).Id.ToString()
                #$app_i_str = ($appointments_al |? {$_.Subject -eq $Subject}).Id.ToString()
                $vacRecActualExt[$j].EventID = $app_i_str

                #$app_i_str = "AAMkAGY3ZjNhMjJiLWZmYTYtNGYxZC1hNjg5LWJiZDMzNzhiMzQxMQBGAAAAAADytyXmrnOhSoNmT78tucJiBwDUkM8khMbcSoayDV7WzxCiAAAAO8DkAAC8DYnfLB5IRouWxsLHtq+aAAIhICRcAAA="

                #$App = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service, $app_i_str)

                #$App = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service_al, $app_i_str)
                $App = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service, $app_i_str)
                $Subject = "Отпуск c " + $appStart + " по " + $appEnd
                $App.Subject = $Subject #+ " " + $mailBoxes[$i].employee_login
                
                $app.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite,`
                [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::SendToNone)
                
                
                $logFile += ",EventCreated"
                $CreatedEventCount += 1
                $vacRecActualExt[$j].dt_last_update = $now_date
                #$vacRecActualExt[$j].IsEventExist = $true
                $vacRecActualExt[$j].IsEventExist = 1
#test ONLY
               # $app.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems,`
               # [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone)

#Inform organizer service. Collect data

                    #$organizer_list = @()
                    $organizer_Userlist = @()
                    $Organizer = $null
                    #$Organizer = ($res_al.appointments_al.Organizer | Sort-Object -Unique)
                    $Organizer = ($res1.appointments.Organizer | Sort-Object -Unique)
                    if( $Organizer -ne $null) {
                        #$mailAddressSelf = Get-Mailbox $myTo | select PrimarySmtpAddress, LegacyExchangeDN
                        $mailAddressSelf = Get-Mailbox $mailboxes[$i].mail | select PrimarySmtpAddress, LegacyExchangeDN 
                    
                        (0..(($Organizer.count)-1)) | %{
                            $param_o = @{
                                organizer_address = $organizer[$_].Address
                                vacation_event = $mailboxes[$i].employee_login + "; " + $mailboxes[$i].employee_nm + "; " + $mailboxes[$i].dt_vacation_from + "; " + $mailboxes[$i].dt_vacation_to }
                         [array]$organizer_Userlist += New-Object -TypeName PSObject -Property $param_o
                        }
                        $organizer_list += $organizer_Userlist | ?{($_.organizer_address -ne $mailAddressSelf.PrimarySmtpAddress) -and ($_.organizer_address -ne $mailAddressSelf.LegacyExchangeDN)}
                    }
                 #$organizer_list.vacation_event -join '<br>'

                  #  $organizer_list |Out-GridView
                  #  $organizer_Userlist |Out-GridView
                #$tr1 = @()
              #  $tr1 += ForEach($org in $organizer_list){(Get-Mailbox $org.organizer_address).PrimarySmtpAddress.Address}
               


         } catch {
            $logFile += ",EventCreationError," + $Error[0]
            $logFileErrors = $true
         }

         try {
               if ($vacRecActualExt[$j].IsEventExist -eq 1) {

                $mySubject = "В ваш почтовый календарь добавлено событие - ОТПУСК"

                $tt1 = $curr_ar_status.InternalMessage

                $body = ("{0}$Subject{1}{6} $oof_def_text {2}" -f $new_event_text_values) +  $tt1 + ("{3}" -f $new_event_text_values) 

                #$myTo = $EmailAddress
                Send-alMessage -body $body -to $EmailAddress -Subj $mySubject
                #Send-alMessage -body $body -to $myTo -Subj $mySubject
                $bodyTmp += '<br>' + $mailBoxes[$i].employee_login + "; " + $Subject
                }

         } catch {
                $logFile += ",EventCreationMessageSendingError," + $Error[0]
             $logFileErrors = $true
         }

        }

      } # end if

    #control autoreply status
         #Write-alError -EntryType Information -Message ("before Set")

#'if user deleted our event' check
         
         $UserDelEvent = $false
         if ($vacRecActualExt[$j].IsEventExist -ne 2 -and $vacRecActualExt[$j].EventID -ne "") {
            if ($appointments -eq $null) { $appointments = (Get-alEWSEvent).Appointments }
            $app_i_str_our = $null
            #$app_i_str = ($appointments |? {$_.Subject -eq $Subject}).Id.ToString()
            if ($appointments -ne $null) {
                $app_i_str_our = $appointments | ?{($_).Id.ToString() -eq $vacRecActualExt[$j].EventID}
            }
            if ($app_i_str_our -eq $null) {
                $vacRecActualExt[$j].IsEventExist = 2
                $logFile += ",$DeletedByUser"
                $DeletedByUserEventCount += 1
            }
         }
         #if user deleted our event. True means the user has deleted our event and we will ommit  OOF control
         if ($vacRecActualExt[$j].IsEventExist -eq 2) {$UserDelEvent = $true}



            #if ($appointments -eq $null) { $appointments = (Get-alEWSEvent).Appointments }

             #if user deleted our event. True means the user has deleted our event and we will ommit  OOF control
         
             #$UserDelEvent = ($appointments | ?{$_.Subject -eq $Subject}) -eq $null -and ($vacRecActualExt[$j].IsEventExist -eq $true)
             #$UserDelEvent = ($appointments | ?{$_.Subject -eq $Subject}) -eq $null -and ($vacRecActualExt[$j].is_deleted -eq 1)
        
       if ((($now_date -ge (Get-Date($mailBoxes[$i].dt_vacation_from))) -and ($now_date -le (Get-Date($mailBoxes[$i].dt_vacation_to)))`
        -and $vacRecActualExt[$j].IsOOFOn -ne 1 -and $vacRecActualExt[$j].IsOOFOn -ne 2) -and $UserDelEvent -eq $false) {
        # -and $vacRecActualExt[$j].IsOOFOn -ne $true) -and $UserDelEvent -eq $false) {
            $logFile += ",OOF_check"
            $OOF_OffCount += 1
            #Write-alError -EntryType Information -Message ("first if")
            try {
                $curr_ar_status = $mailboxes[$i].mail | Get-MailboxAutoReplyConfiguration |select AutoReplyState, StartTime, EndTime, OOFEventSubject, Identity,IsValid, InternalMessage
            
                $StartTimeDelta = ((Get-Date($mailBoxes[$i].dt_vacation_from)).Date - (Get-Date($curr_ar_status.StartTime)).Date).Days
                $EndTimeDelta = ((Get-Date($mailBoxes[$i].dt_vacation_to)).Date - (Get-Date($curr_ar_status.EndTime)).Date).Days

                $isUserOOFScheduled = $StartTimeDelta -ge -2 -and $EndTimeDelta -le 2 -and $curr_ar_status.AutoReplyState -eq "Scheduled"
#check if OOF status equals 'Scheduled' (i.e. stop to control OOF status at all)             
            if ($vacRecActualExt[$j].IsOOFOn -ne 2 -and $isUserOOFScheduled -eq $true ) {
                $vacRecActualExt[$j].IsOOFOn = 2
                $logFile += ",$OOF_OnControlOff"
                $OOF_OnControlOff += 1
            }
#check if OOF status equals 'Enabled' (i.e. stop to control OOF status at all) 
             if ($vacRecActualExt[$j].IsOOFOn -ne 2 -and $curr_ar_status.AutoReplyState -eq "Enabled") { $vacRecActualExt[$j].IsOOFOn = 2 }

           # if (-not((($StartTimeDelta -ge -2 -and $EndTimeDelta -le 2 -and $curr_ar_status.AutoReplyState -eq "Scheduled"))`
           #  -or $curr_ar_status.AutoReplyState -eq "Enabled")) {
             
             if (-not($vacRecActualExt[$j].IsOOFOn -eq 2 -or $curr_ar_status.AutoReplyState -eq "Enabled")) {
                #Write-alError -EntryType Information -Message ("2nd if")
                try {
                $bodyTmp += "<br>" + $mailBoxes[$i].employee_login + "; " + $mailBoxes[$i].mail + "; " + $mailBoxes[$i].dt_vacation_from  + "; " + $mailBoxes[$i].dt_vacation_to + "; " + $mailBoxes[$i].employee_vacation_type_nm
                $bodyTmp += "<br>&emsp;&ensp;" + $curr_ar_status.AutoReplyState + "; " + $curr_ar_status.StartTime + "; " + $curr_ar_status.EndTime 
                #Write-alError -EntryType Information -Message ("Set")
                Set-MailboxAutoReplyConfiguration $mailboxes[$i].mail –InternalMessage $oof_def_text -AutoReplyState Scheduled –StartTime $d_start -EndTime $d_end.AddDays(1) -ExternalAudience None
                #$vacRecActualExt[$j].IsOOFOn = $true
                $vacRecActualExt[$j].IsOOFOn = 1
                $logFile += ",OOF_On"
                $OOF_OnCount += 1
                #test
                #$test_mail = ('Автоответ "Нет на работе" включен (' + $mailboxes[$i].mail + ')')
                #Send-alMessage -body $oof_def_text -to $myTo -Subj $test_mail
                } catch {
                     $logFile += ",SetAutoReplyConfigurationError," + $Error[0]
                     $logFileErrors = $true
                }

#send OOF-message to user
                #send OOF-message to user if auto event has not been created
                #if (-not($vacRecActualExt[$j].IsEventExist -eq $true)) {

                #the next comment means to send OOF-mesage alwyas. Check conditions before uncomment!
                if ($vacRecActualExt[$j].IsOOFOn -eq 1) {
                    try {
                    $mySubject = 'Автоответ "Нет на работе" включен (событие ОТПУСК)'


                    $tt1 = $curr_ar_status.InternalMessage
                    
                    $body = ("{4} $oof_def_text {2}" -f $new_event_text_values) +  $tt1 + ("{3}" -f $new_event_text_values)

                    #$myTo = $EmailAddress
            Send-alMessage -body $body -to $EmailAddress -Subj $mySubject
            #send e-mail one time and stop OOF control
            $vacRecActualExt[$j].IsOOFOn = 2
                    #Send-alMessage -body $body -to $myTo -Subj $mySubject $test_mail
                    $bodyTmp += '<br>&emsp;&ensp;OOF_OnMessageSent:' + $mailBoxes[$i].employee_login + "; " + $Subject
                    $logFile += ",OOF_OnMessageSent"
                    $OOF_OnMessageSentCount += 1
                    } catch {
                        $logFile += ",OOFTurnOnMessageSendingError," + $Error[0]
                        $logFileErrors = $true
                        $vacRecActualExt[$j].IsOOFOn = 2
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
    
#Inform organizer service. Send mail
                $organizer_listUnique = $organizer_list.organizer_address | Sort-Object -Unique #|Out-GridView
                $organizer_listUnique_smtp = @()
                $organizer_listUnique_smtp += ForEach($org in $organizer_listUnique){
                    #(Get-Mailbox $org).PrimarySmtpAddress.Address | select @{n='organizer_address'; e={$org}}, @{ n='smtpAddress'; e={$_}}
                    (Get-Mailbox $org -ErrorAction SilentlyContinue).PrimarySmtpAddress.Address | select @{n='organizer_address'; e={$org}}, @{ n='smtpAddress'; e={$_}}
                    }
                    #$organizer_listUnique_smtp |Out-GridView
                $bodyTmp += '<br>------<br>Vacation records for organizer e-mail'
                foreach ($orgAddress in $organizer_listUnique_smtp )  {
                    [PSObject[]]$orgAddressTemp = @()
                    $orgAddressTemp = $organizer_list | ?{$_.organizer_address -eq $orgAddress.organizer_address}
                    #$bodyOrg = "Добрый день.<br><br>Для Информации.<br><br>"
                    #$bodyOrg += "Вы назначили мероприятия в календаре почтовой системы сотрудникам, которые взяли отпуск (УЗ; ФИО; Начало; Конец):<br><br>"
                    $bodyOrg = "{5}" -f $new_event_text_values
                    $bodyOrg += $orgAddressTemp.vacation_event -join '<br>'
                    $bodyOrg += "{3}" -f $new_event_text_values 
                    #АК test
                    #$KA = "" 
                    #$KA = ""
                    $KA = $orgAddress.smtpAddress
                    if ($orgAddress.smtpAddress -eq $KA) {
                    Send-alMessage -To $orgAddress.smtpAddress -Subj "Автоинформирование об отпусках для организаторов мероприятий" -body $bodyOrg
                    $OrganizerMessageSent += 1
                    $bodyTmp += "<br>" + ($orgAddressTemp.vacation_event).Count + " for " + ($orgAddress.smtpAddress)
                    #Send-alMessage -To $myTo -Subj "Автоинформирование об отпусках для организаторов мероприятий для $KA" -body $bodyOrg
                    #Send-alMessage -To $myTo -Subj "Автоинформирование об отпусках для организаторов мероприятий для $KA" -body $bodyTmp
              }
              }
            $organizer_list | Export-Csv -Path $file_ois -Delimiter ";" -Encoding UTF8 -NoTypeInformation
            #$organizer_list = Import-Csv -Path $file_ois -Delimiter ";" -Encoding UTF8

        
     #Write-alError -EntryType Information -Message ("Before Stop")
    
    $body1 += '<br>' + $MustEventCount + " : Count of events should be exist"
    $body1 += '<br>' + $NoEventCount + " : Count of events that are not exist"
    $body1 += '<br>' + $CreatedEventCount + " : Count of created events"
    $body1 += '<br>' + $DeletedEventCount + " : Count of deleted events"
    $body1 += '<br>' + $UpdatedEventCount + " : Count of updated events"
    $body1 += '<br>' + $NoNeedEventCount + " : Count of events that are not to be necessary created (appropriate user-created events have been found)"
    $body1 += '<br>' + $DeletedByUserEventCount + ": Count of events has been deleted by user"
    $body1 += '<br>' + $OOF_OffCount + " : Count of OOF-Off records should be checked"
    $body1 += '<br>' + $OOF_OnCount + " : Count of OOF turned On"
    $body1 += '<br>' + $OOF_OnMessageSentCount + " : Count of 'OOF turned On' messages have been sent to user"
    $body1 += '<br>' + $OOF_OnControlOff + " : Count of OOF that do not need to be activated (appropriate user-created OOF-schedule have been found)"
    $body1 += '<br>' + $OrganizerMessageSent + " : Count of messages has been sent to meeting organizers"
    
    if($logFileErrors) {
        $body1 += '<br><br> There were some errors. Addition info logfile ' + $file_log
    } 

    
    #Write-alError -EntryType Information -Message ("Stop")

    $vacRecActualExt | %{if ((get-date($_.dt_vacation_to)).Date -ge $now_date ) {$_}} | Export-Csv -Path $file_ext -Delimiter ";" -Encoding UTF8 -NoTypeInformation
    
    #$vacRecActualExt | Export-Csv -Path $file_ext -Delimiter ";" -Encoding UTF8 -NoTypeInformation

    Write-alError -EntryType Information -Message ($body1.Replace('<br>', "`r`n")) -EvenSourceName $EvenSourceName -ApplicationName $ApplicationName

    $body1 += '<br>------<br>' + $bodyTmp

    Send-alMessage -body $body1 -to $myTo -Subj $mySubj
    
    Remove-alLogFiles -log_Path $logPath -logDepth 9 #remove log files older then 9 days
<#
    #OOF-ON verification
    $body1 = ""
    $mySubj = "OOF-ON verification"
    for ($i = 0; $i -lt $mailBoxes.Count; $i++) {
        $curr_ar_status = $mailboxes[$i].mail | Get-MailboxAutoReplyConfiguration |select AutoReplyState, StartTime, EndTime, OOFEventSubject, Identity,IsValid, InternalMessage
       $body1 += '<br>' + $mailboxes[$i].mail + '<br>&emsp;&ensp;' + ($curr_ar_status | %{$_ -join '<br>&emsp;&ensp;'})
       #if ((Get-date($curr_ar_status.StartTime) -Format 'HH') -eq '00' -and $curr_ar_status.AutoReplyState -eq "Scheduled") {
       # $curr_ar_status
        ##Set-MailboxAutoReplyConfiguration $mailboxes[$i].mail -AutoReplyState Disabled
       #}

    }
    Send-alMessage -body $body1 -to $myTo -Subj $mySubj
#>

} catch {
      $vacRecActualExt | Export-Csv -Path $file_ext -Delimiter ";" -Encoding UTF8 -NoTypeInformation
      Out-File -InputObject $logFile -Encoding utf8 -FilePath $file_log
      Write-Host $Error[0].Exception.Message
      $body =  "Error : " + $Error[0].Exception.Message
      $body += "<br>" + ($Error[0].CategoryInfo -join '`r`n')
      $body += "<br>" + $Error[0].ScriptStackTrace
      $body += "<br><br> Script execution fails. Addition info logfile " + $file_log
      Write-alError -EntryType Error -Message $body -EvenSourceName $EvenSourceName -ApplicationName $ApplicationName
      #Write-alError -EntryType Error -Message "Zabbix test5"
      Send-alMessage -body $body -To $myTo -Subj $mySubj
} 
