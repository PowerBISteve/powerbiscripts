
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  

$Form = New-Object System.Windows.Forms.Form    
$Form.Size = New-Object System.Drawing.Size(600,400)  

############################################## Start functions
Function Get-FileName($var, $initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "PBIX (*.pbix)| *.pbix"
    $OpenFileDialog.ShowDialog() | Out-Null
    $var.Text = $OpenFileDialog.filename

}           

############################################## end functions


############################################## Start text fields

$modelName = New-Object System.Windows.Forms.TextBox 
$modelName.Location = New-Object System.Drawing.Size(100,20) 
$modelName.Size = New-Object System.Drawing.Size(400,20) 
$modelName.MultiLine = $False 
$modelName.ScrollBars = "Vertical" 
$modelName.Text = $modelNameValue
$Form.Controls.Add($modelName) 


$reportName = New-Object System.Windows.Forms.RichTextBox 
$reportName.Location = New-Object System.Drawing.Size(100,50) 
$reportName.Size = New-Object System.Drawing.Size(400,20) 
$reportName.MultiLine = $False 
$reportName.ScrollBars = "Vertical" 
$Form.Controls.Add($reportName) 


$scratch = New-Object System.Windows.Forms.RichTextBox 
$scratch.Location = New-Object System.Drawing.Size(100,80) 
$scratch.Size = New-Object System.Drawing.Size(400,20) 
$scratch.text = $env:USERPROFILE + "\Documents\"
$scratch.MultiLine = $False 
$scratch.ScrollBars = "Vertical" 
$Form.Controls.Add($scratch) 


############################################## end text fields

############################################## Start buttons
$Button = New-Object System.Windows.Forms.Button 
$Button.Location = New-Object System.Drawing.Size(20,20) 
$Button.Size = New-Object System.Drawing.Size(75,20) 
$Button.Text = "Get Model" 
$Button.Add_Click({ Get-FileName($modelName) }) 
$Form.Controls.Add($Button) 

$Button = New-Object System.Windows.Forms.Button 
$Button.Location = New-Object System.Drawing.Size(20,50) 
$Button.Size = New-Object System.Drawing.Size(75,20) 
$Button.Text = "Get Report" 
$Button.Add_Click({Get-FileName($reportName) }) 
$Form.Controls.Add($Button) 

$Button = New-Object System.Windows.Forms.Button 
$Button.Location = New-Object System.Drawing.Size(275,120) 
$Button.Size = New-Object System.Drawing.Size(50,30) 
$Button.Text = "GO" 
$Button.Add_Click({Get-FileName($reportName) }) 
$Form.Controls.Add($Button) 

############################################## end buttons
Function Modify-PBIX([string]$inputpath, [string[]]$filestoremove){

    #Make temp folder
    $temppth = $env:TEMP  + "\PBI TEMP"
    If(!(test-path $temppth))
    {New-Item -ItemType Directory -Force -Path $temppth}

    #Unpackage pbix
    $zipfile = ($inputpath).Substring(0,($inputpath).Length-4) + "zip"
    Rename-Item -Path $inputpath -NewName  $zipfile
              
    #Initialise object
    $ShellApp = New-Object -COM 'Shell.Application'
    $InputZipFile = $ShellApp.NameSpace( $zipfile )

    #Move files to temp
    foreach ($fn in $filestoremove){ 
       $InputZipFile.Items() | ? {  ($_.Name -eq $fn) }  | % {
       $ShellApp.NameSpace($temppth).MoveHere($_)   }  
    }
    
    #Delete temp
    Remove-Item ($temppth) -Recurse
    
    #Repackage 
    Rename-Item -Path $zipfile -NewName $inputpath  
}




#Choose file
try {$pathn = Get-FileName}
catch { "Incompatible File" }


#Check for errors
If([string]::IsNullOrEmpty($pathn )){            
    exit } 

elseif ( IsFileLocked($pathn) ){
    exit } 

#Run Script
else{    

    #set variables
    $modelfiles   = @( 'SecurityBindings', 'Report')
    $reportfiles   = @('Connections','DataModel',  'SecurityBindings')
    
    #Copy files
    $pathf = Get-ChildItem $pathn
    $reportname = [io.path]::GetFileNameWithoutExtension($pathn)
    $model = ($pathf).toString().Replace('.pbix', '_model.pbix')
    $report = ($pathf).toString().Replace('.pbix', '_report.pbix')    
    Copy-Item $pathn -Destination $model`
    Copy-Item $pathn -Destination $report

    #modify files
    Modify-PBIX $model $modelfiles
    Modify-PBIX $report $reportfiles
    
}


#$Form.Add_Shown({$Form.Activate()})
#[void] $Form.ShowDialog()

