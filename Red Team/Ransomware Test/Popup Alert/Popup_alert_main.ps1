<# was made as an alert system if the email was down, replace the images with different alert scenarios.#>

function preview { 
    if($combo_box_selected -like "Evacuate") { 
           # [console]::beep(1000,1000) #plays beep sound when form opens. 

[void][reflection.assembly]::LoadWithPartialName("System.Windows.Forms")  
$fire_file = (get-item '\\files\path\folder\evacuate.jpg')

$img2 = [System.Drawing.Image]::Fromfile($fire_file);

[System.Windows.Forms.Application]::EnableVisualStyles();
$form = new-object Windows.Forms.Form
$form.Text = "Attention $env:username, this is an emergency"
$form.Width = $img2.Size.Width;
$form.Height =  $img2.Size.Height;
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Width =  $img2.Size.Width;
$pictureBox.Height =  $img2.Size.Height;
$form.WindowState = "maximized"
$pictureBox.Image = $img2;
$form.controls.add($pictureBox)
$form.Add_Shown( { $form.Activate() } )
$form.ShowDialog()
}
    elseif($combo_box_selected -like "Active Shooter") { 
        [void][reflection.assembly]::LoadWithPartialName("System.Windows.Forms")  
$active_file = (get-item "\\files\path\folder\shooter.jpg")

$img3 = [System.Drawing.Image]::Fromfile($active_file);

[System.Windows.Forms.Application]::EnableVisualStyles();
$form = new-object Windows.Forms.Form
$form.Text = "Attention $env:username, this is an emergency"
$form.Width = $img3.Size.Width;
$form.Height =  $img3.Size.Height;
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Width =  $img3.Size.Width;
$pictureBox.Height =  $img3.Size.Height;
$form.WindowState = "maximized"
$pictureBox.Image = $img3;
$form.controls.add($pictureBox)
$form.Add_Shown( { $form.Activate() } )
$form.ShowDialog()
}
    else { 
        #[console]::beep(1000,1000) #plays beep sound when form opens. 

[void][reflection.assembly]::LoadWithPartialName("System.Windows.Forms")  
$shelter_file = (get-item '\\files\path\folder\shelter.jpg')


$img1 = [System.Drawing.Image]::Fromfile($shelter_file);

[System.Windows.Forms.Application]::EnableVisualStyles();
$form = new-object Windows.Forms.Form
$form.Text = "Attention $env:username, this is an emergency"
$form.Width = $img1.Size.Width;
$form.Height =  $img1.Size.Height;
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Width =  $img1.Size.Width;
$pictureBox.Height =  $img1.Size.Height;
$form.WindowState = "maximized"
$pictureBox.Image = $img1;
$form.controls.add($pictureBox)
$form.Add_Shown( { $form.Activate() } )
$form.ShowDialog()
    }   
}

$i = 0
function Send-Popup_Message { 
    $login = Get-Credential $env:USERNAME #replace with your username.
    $task_block = {
    $path = "\\files\path\folder\Popup_box.ps1"
    $seconds = 1
    <# Use the paramaters from test_popup.ps1 to make sure the variables are passed to the other script. Need the Using: variable scope to 
    make sure you can pass the variable along in a remote communication#>
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument " -WindowStyle hidden -file $path -message $using:combo_box_selected"
    $trigger = New-ScheduledTaskTrigger -Once -At ((Get-Date) + (New-TimeSpan -Seconds $seconds))
    $principal = New-ScheduledTaskPrincipal -UserId (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -expand UserName) 
    $task_name = "!random_task_name_" + (-join ((65..90) + (97..122) | Get-Random -Count 10 | ForEach-Object {[char]$_})) #creates random task name 
    $register = Register-ScheduledTask -TaskName $task_name -Trigger $trigger -Action $action -Principal $principal
    Start-Sleep ($seconds + 1)
    Get-ScheduledTask -TaskName $task_name | Unregister-ScheduledTask -Confirm:$false
    }
    #will put something here to fill array with computers once we decide to launch this. 
    $computers = @("testcomputer-AXXX")
    foreach($computer in $computers) { 
    Invoke-Command -ComputerName $computer -Credential $login -ScriptBlock $task_block 
    $i++
    Write-Progress -Activity "Sending Out Message..." -Status " Sending $i of $($computers.Count)" -CurrentOperation "$computer" -PercentComplete (($i / $computers.count) * 100)
    }

}


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '331,104'
$Form.text                       = "Alert Message"
$Form.TopMost                    = $false
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe") #PSHOME is the defualt directory for powershell (c:\windows\system32)
$form.Icon = $Icon

$combo_box                        = New-Object System.Windows.Forms.ComboBox
$combo_box.text                   = "-Choose-"
$combo_box.width                  = 189
$combo_box.height                 = 54
$combo_box.location               = New-Object System.Drawing.Point(10,14)
$combo_box.Font                   = 'Microsoft Sans Serif,10,style=Bold'
#fill
$fill_combo_box = @("Evacuate", "Shelter in Place","Active Shooter") 
foreach($fill in $fill_combo_box) { 
$combo_box.Items.Add($fill) | Out-Null
}

$Preveiw_Button                  = New-Object system.Windows.Forms.Button
$Preveiw_Button.text             = "Preview"
$Preveiw_Button.width            = 96
$Preveiw_Button.height           = 30
$Preveiw_Button.location         = New-Object System.Drawing.Point(209,18)
$Preveiw_Button.Font             = 'Microsoft Sans Serif,10'

$Send_Button                     = New-Object system.Windows.Forms.Button
$Send_Button.BackColor           = "#ff0000"
$Send_Button.text                = "SEND"
$Send_Button.width               = 97
$Send_Button.height              = 30
$Send_Button.location            = New-Object System.Drawing.Point(208,60)
$Send_Button.Font                = 'Microsoft Sans Serif,10'
$Send_Button.ForeColor           = "#ffffff"

$Form.controls.AddRange(@($combo_box,$Preveiw_Button,$Send_Button))


$Preveiw_Button.Add_click({
    $global:combo_box_selected = $combo_box.SelectedItem.ToString();
    preview
})


$Send_Button.Add_Click({
    $global:combo_box_selected = "`"$($combo_box.SelectedItem.ToString())`"";
    Send-Popup_Message
})



[void]$form.ShowDialog()
Read-Host -Prompt "press any button to continue..."