#A different way to pull all the computers in different OUs so you dont have to type them all out. 
<#$OUs = Get-ADOrganizationalUnit -server $ADServer ` -SearchBase "OU=_Departments,DC=domain,DC=local"  -Filter 'name -eq "computers"'
$DomainCPUs = Get-ADComputer -server $ADServer ` -filter * -searchbase "CN=Computers,DC=domain,DC=local" ` | Select-Object Name
for($i=0;$i -le $OUs.length-1;$i++){
		$CPUs = Get-ADComputer -server $ADServer ` -filter * -searchbase $OUS[$i] ` | Select-Object Name 
        $AllCPUs += $CPUs 
        $AllCPUs += $DomainCPUs 
        $AllCPUs = $AllCPUs | sort-object Name
    }#>
function Preview-MessageBox { 
    if($Type_ComboBox_selected -like "Informational") { 
        Add-Type -AssemblyName System.Windows.Forms 
        $global:balloon = New-Object System.Windows.Forms.NotifyIcon
        $path = (Get-Process -id $pid).Path
        $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
        $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
        $balloon.BalloonTipText = "$input_box_selected"
        $balloon.BalloonTipTitle = "Hello $env:USERNAME"
        $balloon.Visible = $true 
        $balloon.ShowBalloonTip(100000)
} 
elseif ($Type_ComboBox_selected -like "Error") {
Add-Type -AssemblyName System.Windows.Forms 
        $global:balloon = New-Object System.Windows.Forms.NotifyIcon
        $path = (Get-Process -id $pid).Path
        $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
        $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Error
        $balloon.BalloonTipText = "$input_box_selected"
        $balloon.BalloonTipTitle = "Attention $env:USERNAME"
        $balloon.Visible = $true 
        $balloon.ShowBalloonTip(100000)
}
else { #warning
Add-Type -AssemblyName System.Windows.Forms 
        $global:balloon = New-Object System.Windows.Forms.NotifyIcon
        $path = (Get-Process -id $pid).Path
        $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
        $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning
        $balloon.BalloonTipText = "$input_box_selected"
        $balloon.BalloonTipTitle = "Attention $env:USERNAME"
        $balloon.Visible = $true 
        $balloon.ShowBalloonTip(100000)
    }
}

#progress bar settings, will display how long the process will take. 

$i = 0

function Send-Popup_Message { 
    $login = new-object System.Management.Automation.PSCredential -argumentlist $Username_TextBox_selected,$Password_MaskedTextBox_selected
    $task_block = {
    $path = "\\files\path\folder\test_popup.ps1" 
    $seconds = 1
    <# Use the paramaters from test_popup.ps1 to make sure the variables are passed to the other script. Need the Using: variable scope to 
    make sure you can pass the variable along in a remote communication#>
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument " -WindowStyle hidden -file $path -text $Using:input_box_selected -messagetype $Using:type_combobox_selected "
    $trigger = New-ScheduledTaskTrigger -Once -At ((Get-Date) + (New-TimeSpan -Seconds $seconds))
    $principal = New-ScheduledTaskPrincipal -UserId (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -expand UserName) 
    $task_name = "!random_task_name_" + (-join ((65..90) + (97..122) | Get-Random -Count 10 | ForEach-Object {[char]$_})) #creates random task name 
    $register = Register-ScheduledTask -TaskName $task_name -Trigger $trigger -Action $action -Principal $principal
    Start-Sleep ($seconds + 1)
    Get-ScheduledTask -TaskName $task_name | Unregister-ScheduledTask -Confirm:$false
    }
    #fills array with all current HQ computers. 
    $loop = "ou=marketing,ou=Departments,DC=domain,DC=local",
    "ou=admins,ou=IT,DC=domain,DC=local",
    "ou=management,ou=accoutning,DC=domain,DC=local" | ForEach-Object { Get-ADcomputer -filter * -SearchBase $_  | select-object Name}
    #$loop = $AllCPUs
    #remove DRC entries so we do not try and remote into them. 

    foreach($computer in $loop.Name) { 
    Invoke-Command -ComputerName $computer -Credential $login -ScriptBlock $task_block 
    $i++ #counts up each time
    Write-Progress -Activity "Sending Out Message..." -Status " Sending $i of $($loop.Count)" -CurrentOperation "$computer" -PercentComplete (($i / $loop.count) * 100)
    }

}
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '368,190'
$Form.text                       = "Toast Message App"
$Form.TopMost                    = $false
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe") #PSHOME is the defualt directory for powershell (c:\windows\system32)
$form.Icon = $Icon

$Input_Box_Label                 = New-Object system.Windows.Forms.Label
$Input_Box_Label.text            = "Type your message below."
$Input_Box_Label.AutoSize        = $true
$Input_Box_Label.width           = 25
$Input_Box_Label.height          = 10
$Input_Box_Label.location        = New-Object System.Drawing.Point(15,11)
$Input_Box_Label.Font            = 'Microsoft Sans Serif,10'

$input_box                       = New-Object system.Windows.Forms.TextBox
$input_box.multiline             = $true
$input_box.width                 = 237
$input_box.height                = 81
$input_box.location              = New-Object System.Drawing.Point(15,32)
$input_box.Font                  = 'Microsoft Sans Serif,10'

$Username_Label                  = New-Object system.Windows.Forms.Label
$Username_Label.text             = "Username"
$Username_Label.AutoSize         = $true
$Username_Label.width            = 25
$Username_Label.height           = 10
$Username_Label.location         = New-Object System.Drawing.Point(8,126)
$Username_Label.Font             = 'Microsoft Sans Serif,10'

$Username_TextBox                = New-Object system.Windows.Forms.TextBox
$Username_TextBox.multiline      = $false
$Username_TextBox.width          = 100
$Username_TextBox.height         = 20
$Username_TextBox.location       = New-Object System.Drawing.Point(82,124)
$Username_TextBox.Font           = 'Microsoft Sans Serif,10'

$Password_Label                  = New-Object system.Windows.Forms.Label
$Password_Label.text             = "Password"
$Password_Label.AutoSize         = $true
$Password_Label.width            = 25
$Password_Label.height           = 10
$Password_Label.location         = New-Object System.Drawing.Point(8,154)
$Password_Label.Font             = 'Microsoft Sans Serif,10'

$Password_MaskedTextBox          = New-Object system.Windows.Forms.MaskedTextBox
$Password_MaskedTextBox.multiline  = $false
$Password_MaskedTextBox.width    = 100
$Password_MaskedTextBox.height   = 20
$Password_MaskedTextBox.location  = New-Object System.Drawing.Point(82,151)
$Password_MaskedTextBox.Font     = 'Microsoft Sans Serif,10'
$Password_MaskedTextBox.UseSystemPasswordChar = $true

$Type_Label                      = New-Object system.Windows.Forms.Label
$Type_Label.text                 = "Message Type"
$Type_Label.AutoSize             = $true
$Type_Label.width                = 25
$Type_Label.height               = 10
$Type_Label.location             = New-Object System.Drawing.Point(203,127)
$Type_Label.Font                 = 'Microsoft Sans Serif,10'

$Type_ComboBox                   = New-Object system.Windows.Forms.ComboBox
$Type_ComboBox.text              = "-Choose-"
$Type_ComboBox.width             = 100
$Type_ComboBox.height            = 20
$Type_ComboBox.location          = New-Object System.Drawing.Point(201,151)
$Type_ComboBox.Font              = 'Microsoft Sans Serif,10'

<# Combo box fill#>

$type_combo_box_info = @("Informational", "Warning", "Error")
foreach($fill in $type_combo_box_info) { 
    $Type_ComboBox.Items.Add($fill)
}

$Preview_Button                  = New-Object system.Windows.Forms.Button
$Preview_Button.text             = "Preview"
$Preview_Button.width            = 89
$Preview_Button.height           = 23
$Preview_Button.location         = New-Object System.Drawing.Point(264,37)
$Preview_Button.Font             = 'Microsoft Sans Serif,10'

$Send_Button                     = New-Object system.Windows.Forms.Button
$Send_Button.text                = "Send"
$Send_Button.width               = 87
$Send_Button.height              = 30
$Send_Button.location            = New-Object System.Drawing.Point(264,77)
$Send_Button.Font                = 'Microsoft Sans Serif,10'


$Form.controls.AddRange(@($Input_Box_Label,$input_box,$Username_Label,$Username_TextBox,$Password_Label,$Password_MaskedTextBox,$Type_Label,$Type_ComboBox,$Preview_Button,$Send_Button))

<# Event Handlers#>

$Preview_Button.Add_Click({
    $Global:input_box_selected = $input_box.Text;
    $Global:Type_ComboBox_selected = $Type_ComboBox.SelectedItem.ToString();
    Preview-MessageBox

})

$Send_Button.Add_Click({
    <# Alright guy sit down and listen here, cause this is where the magic happens
    you NEED to wrap the text in literal quotes to include spaces in the string for this to work, if you do not it will only take the first word and display nothing else.#>
    $Global:input_box_selected = "`"$($input_box.Text)`"";
    $Global:Type_ComboBox_selected = $Type_ComboBox.SelectedItem.ToString();
    $Global:Username_TextBox_selected = $Username_TextBox.Text;
    $Global:Password_MaskedTextBox_selected = $Password_MaskedTextBox.Text | ConvertTo-SecureString -AsPlainText -Force; #use this to make the password readable for the host
    Send-Popup_Message
})
[void] $form.Showdialog() 

 Read-Host -Prompt "Press any button to exit..."