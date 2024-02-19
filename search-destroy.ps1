<#
    .SYNOPSIS
        Поиск и удаление писем на предприятии
        Запустить этот файл в консоли Powershell ипоявится новый коммандлет Search-Destroy, можно табулятором после знака тире "-" перебирать возможные параметры. 
 
    .DESCRIPTION
        Функция состоит из двух штатных коммандлетов: Get-MessageTrackingLog и Search-Mailbox.
        В самом начале, принимаются сразу все параметры и раскидываются каждые в свой коммандлет.

        MessageTrackingLog находит поиском всех получателей, которым было сообщение доставлено и передает их 
        на Search-Mailbox, который выполняет необходимые действия. 
 
    .PARAMETER 'MessageSubject', 'Start', 'EventId', 'ResultSize','Sender'
        Штатные параметры ввода в Get-MessageTrackingLog
 
    .PARAMETER 'EstimateResultOnly','DeleteContent','LogOnly','Force','LogLevel','TargetMailbox','TargetFolder'
        Штатные параметры ввода в Search-Mailbox
 
    .EXAMPLE
        Выведет количество писем по ящикам в котортые было доставлено сообщение с темой "Очень тестовое!" от oaaulov@rttv.ru:
        Search-Destroy -Sender oaaulov@rttv.ru -MessageSubject "Очень тестовое!" -Start 01.15.2024 -EstimateResultOnly

        Отправит подробный отчет админу в указанную папку по результату поиска в ящиках в которые было доставлено сообщение:
        Search-Destroy -Sender oaaulov@rttv.ru -MessageSubject "Очень тестовое!" -TargetMailbox avpesterev@rttv.ru -TargetFolder Test -LogLevel Full -LogOnly 
        
        Найдет и удалит все найденные письма, отправит копии в ящик админа
        Search-Destroy -Sender oaaulov@rttv.ru -MessageSubject "Очень тестовое!" -TargetMailbox avpesterev@rttv.ru -TargetFolder Test -DeleteContent -Force

        Просто удалит без возможности восстановления
        Search-Destroy -Sender oaaulov@rttv.ru -MessageSubject "Очень тестовое!" -DeleteContent -Force 
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
        $exch= (Get-ADComputer -Filter "name -like 's-ex-0*'").name| Get-Random
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
        write-host There is no result with options:  @Track -BackgroundColor Red -ForegroundColor White
    }
}