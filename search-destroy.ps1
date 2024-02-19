<#
    .SYNOPSIS
        Search and delete messages in your Exchange Enterprise
 
    .DESCRIPTION
        This function combines two standart cmdlets: Get-MessageTrackingLog and Search-Mailbox, and accepts all 
        the parameters necessary for searching and deleting requested messages.

        MessageTrackingLog finds message recipients and forwards them to Search-Mailbox that performs the necessary 
        actions in the mailboxes

        This only Line, that you should to correct:
        $exch= (Get-ADComputer -Filter "name -like 's-ex-0*'").name| Get-Random
 
    .PARAMETER 'MessageSubject', 'Start', 'EventId', 'ResultSize','Sender'
        Usual parameters for Get-MessageTrackingLog
 
    .PARAMETER 'EstimateResultOnly','DeleteContent','LogOnly','Force','LogLevel','TargetMailbox','TargetFolder'
        Usual parameters for Search-Mailbox
 
    .EXAMPLE
        Shows amount of messages with "Very Test!" subject from some@domain.com:
        Search-Destroy -Sender some@domain.com -MessageSubject "Very Test!" -Start 01.11.2024 -EstimateResultOnly

        Sends full report in discovery mailbox Test Folder, who and when received a letter from such a sender and with such a subject:
        Search-Destroy -Sender some@domain.com -MessageSubject "Very Test!" -TargetMailbox discovery@domain.com -TargetFolder Test -LogLevel Full -LogOnly 
        
        Finds and deletes all found messages, sends copies to the discovery mailbox in Test folder:
        Search-Destroy -Sender some@domain.com -MessageSubject "Very Test!" -TargetMailbox discovery@domain.com -TargetFolder Test -DeleteContent -Force  
#>


function global:Search-Destroy {
    param (
        [Parameter(Mandatory=$true)]$MessageSubject,
        $Sender,
        $Start,
        $EventId="deliver",
        $ResultSize="unlimited",
        $TargetMailbox,
        $TargetFolder,
        [ValidateSet("Full")]$LogLevel,
        [switch]$EstimateResultOnly,
        [switch]$DeleteContent,
        [switch]$LogOnly,
        [switch]$Force
    )
    while (-Not(Get-PSSession|Where-Object ConfigurationName -eq "Microsoft.Exchange")) {
        $exch= (Get-ADComputer -Filter "name -like 's-ex-0*'").name| Get-Random  #Here you need to correct the names of mail servers to suit your template
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$exch/Powershell" -Authentication Kerberos
        Import-PSSession $session -DisableNameChecking -AllowClobber | out-null
        Write-Host Exchange PSSession loaded successfully! -ForegroundColor magenta
    }

    $Track =@{}
    $Search=@{}
    'MessageSubject', 'Start', 'EventId', 'ResultSize','Sender' |Foreach{ 
        $value = if ($PSBoundParameters.ContainsKey($_)) { $PSBoundParameters[$_] } 
         else { Get-Variable -Scope Local -ErrorAction Ignore -ValueOnly $_ }
         if($value) {$Track.Add($_, $value) }
    }

    'EstimateResultOnly','DeleteContent','LogOnly','Force','LogLevel','TargetMailbox','TargetFolder' |Foreach{ 
        $value = if ($PSBoundParameters.ContainsKey($_)) {$PSBoundParameters[$_] } 
         else { Get-Variable -Scope Local -ErrorAction Ignore -ValueOnly $_ }
         if($value) {$Search.Add($_, $value) }
    }

    $users=(Get-TransportService).name |Get-MessageTrackingLog @Track| Select-Object -Unique -ExpandProperty Recipients
    if ($users) {
        $Query = "Subject:$MessageSubject"
        if($Sender){
            $Query += " AND from:$sender"
        }
        if($start){
            $Query += " AND received>=$Start"
        }
            $users | foreach {Search-Mailbox $_ -SearchQuery $Query @Search -WarningAction "SilentlyContinue"} | Format-Table -AutoSize -Wrap @{n="Name";e={$_.Identity -replace '.*/'}},ResultItemsCount
    }
    else {
        write-host There is no result with options: @Track -BackgroundColor Red -ForegroundColor White
    }
}
