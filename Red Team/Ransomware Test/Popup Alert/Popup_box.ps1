param ( 
    [string]$message
)

if($message -like "Evacuate") { 
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
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe") #PSHOME is the defualt directory for powershell (c:\windows\system32)
$form.Icon = $Icon
$pictureBox.Image = $img2;
$form.controls.add($pictureBox)
$form.Add_Shown( { $form.Activate() } )
$form.ShowDialog()
}
elseif($message -like "Shelter in Place") { 
 [void][reflection.assembly]::LoadWithPartialName("System.Windows.Forms")  
$active_file = (get-item "\\files\path\folder\shelter.jpg")

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
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe") #PSHOME is the defualt directory for powershell (c:\windows\system32)
$form.Icon = $Icon
$pictureBox.Image = $img3;
$form.controls.add($pictureBox)
$form.Add_Shown( { $form.Activate() } )
$form.ShowDialog()
}
else { 
 #[console]::beep(1000,1000) #plays beep sound when form opens. 

[void][reflection.assembly]::LoadWithPartialName("System.Windows.Forms")  
$shelter_file = (get-item '\\files\path\folder\shooter.jpg')


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
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe") #PSHOME is the defualt directory for powershell (c:\windows\system32)
$form.Icon = $Icon
$pictureBox.Image = $img1;
$form.controls.add($pictureBox)
$form.Add_Shown( { $form.Activate() } )
$form.ShowDialog()
}   