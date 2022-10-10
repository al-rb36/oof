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

try {

#get current vacation data
    
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
    #$memb_new_SamAccountName_list1 |select -First 5
    #$file1 = $env:USERPROFILE +"\Documents\add_OOF_psscript\20220908-190034.csv"
    $file = $env:USERPROFILE +"\Documents\add_OOF_psscript\vacation_list.csv"
    
    $tmp2 = ""
    if (Test-Path -Path $file -PathType Leaf) {
        $tmp2 = Import-Csv -Path $file -Delimiter ";" -Encoding UTF8 #|select *,  @{n='mail'; e={""}}, @{n='isMailBox'; e={""}}
        }
    #$tmp2 | select -First 3
    #$tmp2 |Get-Member
    #$tmp2.Count
    #$diff_email_list = Compare-Object -ReferenceObject $tmp2 -DifferenceObject $memb_new_SamAccountName_list1 -IncludeEqual # -Property SAMAccountNameAD,PregLeaveName,PregLeaveFrom,PregLeaveTo
    $diff_email_list = $memb_new_SamAccountName_list1 | %{if(-not($tmp2.FullString -match $_.FullString)){$_}}
    $vacRecAdd = $diff_email_list
    #$diff_email_list |Get-Member
    #$diff_email_list |select -First 3 SideIndicator
    #$vacRecAdd = $diff_email_list | ?{$_.SideIndicator -eq "<=" }
    #$vacRecAdd = $diff_email_list

    #$vacRecAdd = $memb_new_SamAccountName_list1 | %{if(-not($tmp2.SAMAccountNameAD -match $_.SAMAccountNameAD)){$_}}
    #$vacRecHold = $memb_new_SamAccountName_list1 | %{if($tmp2.FullString -match $_.FullString){$_}}
    #$vacRecUpdate = $tmp2 | %{if(($memb_new_SamAccountName_list1.SAMAccountNameAD -match $_.SAMAccountNameAD)`
    #     -and ($memb_new_SamAccountName_list1.PregLeaveFrom -lt $_.PregLeaveFrom) ) {$_}}
    #$vacRecOld = $diff_email_list | ?{$_.SideIndicator -ne "<=" }
    #$vacRecAdd.Count | Select -First 3
    #$vacRecOld.Count
    
    $body = [string]$vacRecAdd.Count + " : Count of new vacation records to add<br>------<br>Added vacation records<br>------<br>"
    #$body += ($vacRecAdd |Out-String).Split('`r`n')  -join '<br>'
    $body += $vacRecAdd.FullString  -join '<br>'
    
    if ($vacRecAdd -ne $null) {$tmp2 += $vacRecAdd}
    
    $now_date = (Get-date).Date
    $vacRecActual = $tmp2 | ?{((Get-Date($_.PregLeaveTo) -ErrorAction SilentlyContinue).Date) -ge $now_date}
    $body += "<br>" + $vacRecActual.Count + " : Count of actual vacation records"
    $vacRecDel = $tmp2 | ?{((Get-Date($_.PregLeaveTo) -ErrorAction SilentlyContinue).Date) -lt $now_date.AddDays(0)}
    $body += "<br>" + $vacRecDel.Count + " : Count of vacation records to delete<br>------<br>Deleted vacation records<br>------<br>"
    #$body += ($vacRecDel |Out-String).Split('`r`n') -join '<br>'
    $body += $vacRecDel.FullString -join '<br>'
    $vacRecActual | Export-Csv -Path $file -Delimiter ";" -Encoding UTF8
    #$tmp2 |select SAMAccountNameAD,PregLeaveName,PregLeaveFrom,PregLeaveTo, @{n='StartShift'; e={0}}, @{n='EndShift'; e={0}}, @{n='IsEventExist'; e={$false}}, @{n='IsOOFOn'; e={$false}}, FullString | Export-Csv -Path $file -Delimiter ";" -Encoding UTF8
    #$tmp2 | Export-Csv -Path $file -Delimiter ";" -Encoding UTF8
    #$diff = $rb_raw_list20 | %{if(-not($rb_raw_list10.FullString -match $_.FullString)){$_}}


    
    #$now_date = (Get-date).Date
    #$vacRecActual = $memb_new_SamAccountName_list1 | ?{((Get-Date($_.PregLeaveTo) -ErrorAction SilentlyContinue).Date) -ge $now_date}
    #$body = "<br>" + $vacRecActual.Count + " : Count of actual vacation records"
    #$vacRecActual | Export-Csv -Path $file -Delimiter ";" -Encoding UTF8
    Send-alMessage -body $body

    #$vacRecAdd.InputObject | Select -First 3
    #$tmp2 += $vacRecold.InputObject
    #$vacRecActual.Count

    #$vacRecOld.InputObject | Select -First 9

} catch {
      Write-Host $Error[0].Exception.Message
      $body =  "Error : " + $Error[0].Exception.Message
      $body += "<br>" + ($Error[0].CategoryInfo -join '`r`n')
      $body += "<br>" + $Error[0].ScriptStackTrace
      $body += "<br><br> Script execution fails. No changes were made."
      #Write-alError -EntryType Error -Message $body
      Send-alMessage -body $body
} 