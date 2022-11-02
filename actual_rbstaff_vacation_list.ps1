[Reflection.Assembly]::LoadFile("C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll") | Out-Null
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
#Init data
    $PS_Script_Tilte = ""
    $log_msg = @{}
    $c_in =""
    $domain_rb = ""
    $domain_gts = ""
    $domain_gts_SearchBase = ""
    $connectionString = ""

    [int]$waitTimer = 5 #will try to execute the script Value-number times with 10 minute interval 

    #Send-myMessage constants
    $myFrom = ""
    $mySubject = ""
    $mySmtpServer = ""
    $myTo = ""

    #load init parameters
    ."C:\Scripts\al_PSScripts\OOF_init.ps1"
    
    

    

Function Send-alMessage {
    param (
        $body
    )
    $my_error = ""
    $body += $body_info
    Send-MailMessage -From $myFrom -Subject $mySubject -SmtpServer $mySmtpServer -Body $body -To $myTo -Encoding unicode -BodyAsHtml -ErrorVariable my_error -ErrorAction SilentlyContinue
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

Function Get-alStringHash {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $myStr,
        [Parameter(Mandatory=$true, Position=1)]
        [string] $Algorithm
    )
   
    $mystream = [IO.MemoryStream]::new([byte[]][char[]]([System.Text.Encoding]::UTF8).GetBytes($myStr))
    Get-FileHash -InputStream $mystream -Algorithm $Algorithm
}

for ( $k = 0; $k -lt $waitTimer; $k++) {

try {

#get current vacation data
    $body = ""
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
    $memb_new_SamAccountName_list1 = $memb_new_SamAccountName_list |select SAMAccountNameAD,PregLeaveName,PregLeaveFrom,PregLeaveTo, @{n='StartShift'; e={0}}, @{n='EndShift'; e={0}}, @{n='IsEventExist'; e={$false}}, @{n='IsOOFOn'; e={$false}}, @{n='FullString'; e={$_.SAMAccountNameAD + $_.PregLeaveName + $_.PregLeaveFrom + $_.PregLeaveTo}} #, @{n='mail'; e={""}}, @{n='isMailBox'; e={""}}
    
    $file = $env:USERPROFILE +"\Documents\add_OOF_psscript\vacation_list.csv"
    
    $tmp2 = ""
    if (Test-Path -Path $file -PathType Leaf) {
        $tmp2 = Import-Csv -Path $file -Delimiter ";" -Encoding UTF8 #|select *,  @{n='mail'; e={""}}, @{n='isMailBox'; e={""}}
        }
   
    $diff_email_list = $memb_new_SamAccountName_list1 | %{if(-not($tmp2.FullString -match $_.FullString)){$_}}
    $vacRecAdd = $diff_email_list
    
    $body = [string]$vacRecAdd.Count + " : Count of new vacation records to add"
    #$body += ($vacRecAdd |Out-String).Split('`r`n')  -join '<br>'
    
    
    if ($vacRecAdd -ne $null) {$tmp2 += $vacRecAdd}
    
    $now_date = (Get-date).Date
    $vacRecActual = $tmp2 | ?{((Get-Date($_.PregLeaveTo) -ErrorAction SilentlyContinue).Date) -ge $now_date}
    $body += "<br>" + $vacRecActual.Count + " : Count of actual vacation records"
    $vacRecDel = $tmp2 | ?{((Get-Date($_.PregLeaveTo) -ErrorAction SilentlyContinue).Date) -lt $now_date.AddDays(0)}
    $body += "<br>" + $vacRecDel.Count + " : Count of vacation records to delete"
    #$body += ($vacRecDel |Out-String).Split('`r`n') -join '<br>'
    
    $vacRecActual | Export-Csv -Path $file -Delimiter ";" -Encoding UTF8
    #Write-alError -EntryType Information -Message $body
    Write-alError -EntryType Information -Message ($body.Replace('<br>', "`r`n"))
    
    $body += "<br>------<br>Added vacation records<br>------<br>"
    $body += $vacRecAdd.FullString  -join '<br>' #New vacation records had been added
    $body += "<br>------<br>Deleted vacation records<br>------<br>"
    $body += $vacRecDel.FullString -join '<br>' #Old vacation records had been deleted

    Send-alMessage -body $body

    #Success. No further trys needed
    $k = $waitTimer

} catch {
      Write-Host $Error[0].Exception.Message
      $body += "<br><br>Error : " + $Error[0].Exception.Message
      $body += "<br>" + ($Error[0].CategoryInfo -join '`r`n')
      $body += "<br>" + $Error[0].ScriptStackTrace
      $body += "<br><br> Script execution fails. Try " + [string]([int]$k+1) + " of " + $waitTimer
      Write-alError -EntryType Warning -Message $body
      Send-alMessage -body $body
      #Failure. Let's try one more time
      sleep 600
} 
    
}
