#This script was designed by Steve Campbell and provided by PowerBI.tips
#BE WARNED this will alter Power BI files so please make sure you know what you are doing, and always back up your files!
#This is not supported by Microsoft and changes to future file structures could cause this code to break

#--------------- Released 6/2/2020 ---------------
#--- By Steve Campbell provided by PowerBI.tips ---


#Choose pbix funtion
Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "PBIX (*.pbix)| *.pbix"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}


#Error check function
function IsFileLocked([string]$filePath){
    Rename-Item $filePath $filePath -ErrorVariable errs -ErrorAction SilentlyContinue
    return ($errs.Count -ne 0)
}


#Function to Modify files
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
       $ShellApp.NameSpace($temppth).MoveHere($_)  }
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
    Copy-Item $pathn -Destination $model
    Copy-Item $pathn -Destination $report

    #modify files
    Modify-PBIX $model $modelfiles
    Modify-PBIX $report $reportfiles
    
}


