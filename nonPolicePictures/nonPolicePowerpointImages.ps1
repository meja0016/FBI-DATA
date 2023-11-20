 [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "FBI"
$Form.Size = New-Object System.Drawing.Size(200,200)

$l1 = New-Object System.Windows.Forms.label
$l1.Location = New-Object System.Drawing.Size(0,12)
$l1.Size = New-Object System.Drawing.Size(190,13)
$l1.BackColor = "Red"
$l1.ForeColor = "white"
$l1.Text = "                     development"
$Form.Controls.Add($l1)

$l2 = New-Object System.Windows.Forms.label
$l2.Location = New-Object System.Drawing.Size(0,0)
$l2.Size = New-Object System.Drawing.Size(190,15)
$l2.BackColor = "Black"
$l2.ForeColor = "white"
$l2.Text = "                            FBI"
$Form.Controls.Add($l2)

$l2 = New-Object System.Windows.Forms.label
$l2.Location = New-Object System.Drawing.Size(0,25)
$l2.Size = New-Object System.Drawing.Size(190,15)
$l2.BackColor = "Transparent"
$l2.ForeColor = "Black"
$l2.Text = "Web Service - Powerpoint Files (Images and Slides)"
$Form.Controls.Add($l2)


$tb1 = New-Object System.Windows.Forms.TextBox
$tb1.Text = Get-Date -Format "MM-dd-yy"
$tb1.Location = New-Object System.Drawing.Point(20,40)  
$tb1.Size = New-Object System.Drawing.Size(120,23)  
$Form.Controls.Add($tb1)  


    $Button2 = New-Object System.Windows.Forms.Button
    $Button2.Location = New-Object System.Drawing.Size(20,129)
    $Button2.Size = New-Object System.Drawing.Size(135,23)
    $Button2.Text = "Start"

    $Button2.Add_Click(
        {   
        $folders = Get-ChildItem $tb1.Text -Filter *.pptx
        
        foreach($folder in $folders){
            $t = [io.path]::GetFileNameWithoutExtension($folder);
            echo $t
            mkdir -force  "$t"
            mkdir -force  "$t/css"
            mkdir -force  "$t/css/PNG"
            mkdir -force nonPolicePictures 

            cp "Library\$t.pptx" "$t\$t.tmp.zip"
            echo "copy Libary\$t.pptx $t\$t.tmp.zip" 
            Expand-Archive -Path "$t\$t.tmp.zip" -DestinationPath $t\$t.tmp
            mv "$t\$t.tmp\ppt\media\*" "$t\css\PNG"

            $items = Get-ChildItem "$t\css\PNG\*"
            $items | Rename-Item -NewName { $t+"-"+$_.Name };
            
            del -Force -Recurse "$t\$t.tmp"
            del -Force -Recurse "$t\$t.tmp.zip"
            dir "$t/css/PNG" | Select-Object {"insert into nonPolicePictures (name) values ('$_')" } >> "nonPolicePictures/$t-nonPolicePictures.sql"
            dir "$t/css/PNG" | Select-Object {"{'picture','PowerPoint','Picture','"+$_.Name+"','"+$t+"'}" } >> "nonPolicePictures/$t-nonPolicePicturesTaxonomy.sql"
            cp "$t/css/PNG/*" nonPolicePictures
            del -Force -Recurse "$t"
         }
        }
 
    )
    $Form.Controls.Add($Button2)

    $form.ShowDialog()