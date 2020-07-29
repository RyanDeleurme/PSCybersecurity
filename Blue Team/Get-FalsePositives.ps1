
#test if computer is already connected to outlook online
if($null -ne (Get-PSSession | Where-Object { $_.ConfigurationName -like "*Exchange*"})) #this verifys if the connection is successful/did not work. 
{ 
    Write-Host "Connected to exchange..."
}
else 
{
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking
}

$falseP = @()

#will get addresses that were spoofed with our users account from the last week till the current date
$msg = Get-QuarantineMessage -StartReceivedDate (Get-date).AddDays(-7) -EndReceivedDate (get-date) | `
Where-Object {$_.SenderAddress -like "*myCompanyDomain.com"} | Select-Object MessageId, Identity, ReceivedTime, Type, SenderAddress, Subject, Expires

#add to the array
$falseP += $msg

return $falseP