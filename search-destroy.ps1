function global:Search-Destroy {
    param (
        $MessageSubject,
        $Sender,
        [string]$Start,
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
    $ErrorActionPreference="stop"
    while (-Not(Get-PSSession|Where-Object ConfigurationName -eq "Microsoft.Exchange")) {
        $exch= (Get-ADComputer -Filter "name -like 's-ex-0*'").name| Get-Random
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$exch/Powershell" -Authentication Kerberos
        Import-PSSession $session -DisableNameChecking -AllowClobber | out-null
    }

    if($Start){
        try{
            $MyDate=$Start #[datetime]::parseexact($start,'MM.dd.yyyy', $null).ToString('dd/MM/yyyy')
        }
        catch {
            Write-Host "Provide $Start Date as MM.dd.yyyy" -ForegroundColor Red
            break
        }
        #$Date=[DateTime]::ParseExact($start, "MM/dd/yyyy", $null)
        #$MyDate=$Date.ToString("dd\/MM\/yyyy")
    }
    
    $Track =@{}
    $Search=@{}
    'MessageSubject', 'Start', 'EventId', 'ResultSize','Sender' |Foreach { 
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
        if ($MessageSubject -match "&|/|\|*|^|%|$|#|:|{|}"){
            $MessageSubject=$MessageSubject -replace '[^a-zA-Z0-9\p{L}\s\-\.,:_]' -replace '\s+',' '
        }
        if($MessageSubject){$Query = "Subject:$MessageSubject"}
        if($Sender -and $MessageSubject){$Query += " AND from:$sender"}
        if($Sender -and !$MessageSubject){$Query += "from:$sender"}
        if($start){$Query += " AND received>=$MyDate"}
        $users | foreach {Search-Mailbox $_ -SearchQuery $Query @Search -WarningAction "SilentlyContinue"} | Format-Table -AutoSize -Wrap @{n="Name";e={$_.Identity -replace '.*/'}},ResultItemsCount
    }
    else {
        Write-Host There is no result with options:  @Track -BackgroundColor Red -ForegroundColor White
    }
}
