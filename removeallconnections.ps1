#This script was designed by Steve Campbell and provided by PowerBI.tips
#BE WARNED this will alter Power BI files so please make sure you know what you are doing, and always back up your files!
#This is not supported by Microsoft and changes to future file structures could cause this code to break

#--------------- Released 5/28/2020 ---------------
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

    #copy Files
    $pathf = Get-ChildItem $pathn
    $reportname = [io.path]::GetFileNameWithoutExtension($pathn)
    $model = ($pathf).toString().Replace('.pbix', '_model.pbix')
    $report = ($pathf).toString().Replace('.pbix', '_report.pbix')    
    Copy-Item $pathn -Destination $model
    Copy-Item $pathn -Destination $report

    #set variables
    $reportfiles   = ('Connections','DataModel',  'SecurityBindings')


    #Unpackage pbix
    [Reflection.Assembly]::LoadWithPartialName('System.IO.Compression')   
    $zipfile = ($report).Substring(0,($report).Length-4) + "zip"
    Rename-Item -Path $report -NewName  $zipfile

    #Delete files
    
    $stream = New-Object IO.FileStream($zipfile, [IO.FileMode]::Open)
    $mode   = [IO.Compression.ZipArchiveMode]::Update
    $zip    = New-Object IO.Compression.ZipArchive($stream, $mode)
    ($zip.Entries | ? { $reportfiles -contains $_.Name }) | % { $_.Delete() }

    #Close zip
    $zip.Dispose()
    $stream.Close()
    $stream.Dispose()

    #Repackage and open
    Rename-Item -Path $zipfile -NewName $report 

}