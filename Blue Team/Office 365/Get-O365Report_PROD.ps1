
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


$end = (get-date).ToShortDateString()
$start = (get-date).AddDays(-7).ToShortDateString()
#large report for lots of events
$detailedReport = Get-MailDetailATPReport  -StartDate $start -EndDate $end | Select-Object Eventtype,date,recipientaddress,senderaddress,subject,Action,MessageTraceId 
$topdetail = $detailedReport | Group-Object -property recipientaddress | Sort-Object -Property count -Descending | Select-Object name,count -First 1
$topattackdetail = $detailedReport | Group-Object -Property Eventtype
$topattackdetailtoast = $detailedReport | Group-Object -Property Eventtype | Sort-Object -Property Count -Descending | Select-Object Name -First 1

# overall traffic for the week, take out successful messages
$overallTraffic = get-mailTrafficatpreport -StartDate $start -EndDate $end | Where-Object {$_.Eventtype -ne 'Message Passed'} | Select-Object EventType,Date,MessageCount 

$overallTraffic | ForEach-Object { 
    $_.date = $_.date.toshortdatestring()
    $_.EventType = $_.date + " " + $_.EventType.tostring()
}
#top spam users for the week
$mailTraffic = Get-MailTrafficTopReport -startdate $start -EndDate $end | Where-Object{$_.Eventtype -ne 'TopMailUser'} | Select-Object Name,date,Eventtype,Messagecount
$mailTraffic | ForEach-Object { 
    $_.date = $_.date.toshortdatestring()
}
$topmail = $mailTraffic | Sort-Object -Property Messagecount -Descending | Select-Object name -first 1
$usermailtraffic = $mailTraffic | Group-Object -Property date 
$usermailtraffic | Add-Member -MemberType aliasproperty -name messagecount -Value Count #need to make alias for count
$title = "Office 365 Security Report" + " " + (get-date).ToShortDateString()
$filepath = "\\server\OtherFolder\Folder\Reports"
$filename = "O365report" + (get-date).ToShortDateString().Replace("/",".")


#building the HTML

New-HTML {
    New-HTMLHeading -HeadingText $title -Heading h3 -Underline
    New-HTMLContent { 
        New-HTMLPanel { 
            New-HTMLToast -IconRegular bell -TextHeader "Most targeted user for phishing:" -Text $topdetail.Name -IconColor Red -BarColorLeft Red -TextColor Black
        }
        New-HTMLPanel { 
            New-HTMLToast -IconColor Red -IconRegular bell -TextColor Black -TextHeader "Most seen vulnerability:" -Text $topattackdetailtoast.name -BarColorLeft Red
        }
        New-HTMLPanel { 
            New-HTMLToast -IconColor Red -IconRegular bell -TextColor Black -TextHeader "Most targeted user for spam:" -Text $topmail.name -BarColorLeft Red
        }
    }
    New-HTMLContent -HeaderText "Detailed Threat Report Office 365" -CanCollapse { 
        New-HTMLPanel { 
            New-HTMLChart -Title "Threats by day" -TitleAlignment center { 
                New-ChartBarOptions -Type barStacked 
                New-ChartLegend -Names 'Vulnerabilities' 
                foreach($t in $topattackdetail) { 
                    New-ChartBar -Name $t.name -Value $t.count
                }
            }
        }
        New-HTMLPanel { 
            New-HTMLTable -ArrayOfObjects $detailedReport -ScrollY -Filtering -DisablePaging { 
                New-HTMLTableRowGrouping -Name 'EventType'
            }
        }
        New-HTMLPanel { 
            New-HTMLTable -ArrayOfObjects $detailedReport -ScrollY -Filtering -DisablePaging { 
                New-HTMLTableRowGrouping -Name 'Recipientaddress'
            }
        }
    }
    New-HTMLContent -HeaderText "Vulnerable by the Week" -CanCollapse { 
        New-HTMLPanel { 
            New-HTMLChart -Title "Vulnerable Users" -TitleAlignment center { 
                New-ChartAxisX -Names $usermailtraffic.Name -TitleText "Week"
                New-ChartLine -Name "Spam and phishing count" -Value $usermailtraffic.messagecount
            }
        }
        New-HTMLPanel { 
            New-HTMLChart { 
                New-ChartBarOptions -Type barStacked 
                New-ChartLegend -Names "vulnerabilities during the week"
                foreach($o in $overallTraffic) { 
                    New-ChartBar -Name $o.eventtype -Value $o.MessageCount
                }
            }
        }
    }
} -showhtml -FilePath "$filepath\$filename.html"

#email settings

$Head = 
@"
<style>
  body {
    font-family: "Arial";
    font-size: 10pt;
    color: #4C607B;
    }
  p { 
    font-family: "Arial";
    font-size: 10pt;
    color: #000000;
  }
  .even { background-color: #ffffff; }
  .odd { background-color: #bfbfbf; }
</style>
"@

$mailbody = 
@"
<font color=`"Black`"><h4>Office 365 Phishing Report</h4></font>
<p> Double click on the HTML file to open the Office 365 security report. </br>
This is an automated message, please see ** IT ** for any errors.
</p> 
"@

$post = 
@"
<font color=`"Black`"><h4>Thanks, </h4></font>
<p>John Doe</br>
123.456.7890</br>
<i>IT Guru</i>
</p>
<img src="companylogo.jpg" alt="Logo" width="200" height="80">
"@

$username = "Johnd@example.com"
$password = Read-Host "password?" -AsSecureString
$creds = New-Object System.Management.Automation.PSCredential ($username, $password)
[string]$body = ConvertTo-Html -Head $Head -Body $mailbody -PostContent $post
$mailparams = @{ 
    To="thing@example.com"
    From="johnd@example.com"
    Subject="Weekly O365 Report"
    SmtpServer="smtp.office365.com"
    Body=$body
    BodyasHTML=$true
    UseSSL=$true
    Port=587
    Credential=$creds
    }
Send-MailMessage @mailparams