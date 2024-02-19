# DESCRIPTION
This function combines two standart cmdlets: Get-MessageTrackingLog and Search-Mailbox, and accepts all 
the parameters necessary for searching and deleting requested messages.

MessageTrackingLog finds message recipients and forwards them to Search-Mailbox that performs the necessary 
actions in the mailboxes

This only Line, that you should to correct:<br>
$exch= (Get-ADComputer -Filter "name -like 's-ex-0*'").name| Get-Random
 
## PARAMETER 'MessageSubject', 'Start', 'EventId', 'ResultSize','Sender'
Usual parameters for Get-MessageTrackingLog
 
## PARAMETER 'EstimateResultOnly','DeleteContent','LogOnly','Force','LogLevel','TargetMailbox','TargetFolder'
Usual parameters for Search-Mailbox
 
## EXAMPLE
Shows amount of messages with "Very Test!" subject from some@domain.com:
```PowerShell
Search-Destroy -Sender some@domain.com -MessageSubject "Very Test!" -Start 01.11.2024 -EstimateResultOnly
```
Sends full report in discovery mailbox Test Folder, who and when received a letter from such a sender and with such a subject:
```PowerShell
Search-Destroy -Sender some@domain.com -MessageSubject "Very Test!" -TargetMailbox discovery@domain.com -TargetFolder Test -LogLevel Full -LogOnly 
```
Finds and deletes all found messages, sends copies to the discovery mailbox in Test folder:
```PowerShell
Search-Destroy -Sender some@domain.com -MessageSubject "Very Test!" -TargetMailbox discovery@domain.com -TargetFolder Test -DeleteContent -Force  
```
