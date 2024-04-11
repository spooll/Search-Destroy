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
