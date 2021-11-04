### You will be prompted for these values if not set here
$scadPath = ""
$OutputFolder = ""

################ START OF SCRIPT

Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
function pause{ $null = Read-Host 'Press Any Key or Enter to continue...' }

##################

If ($scadPath -eq ''){

    cls
    echo "You will now be prompted for the SCAD file"
    pause

    $FileBrowser.filter = "scad (*.scad)| *.scad"
    [void]$FileBrowser.ShowDialog()

    $scadPath = $FileBrowser.FileName
}

##################

If ($OutputFolder -eq ''){

    cls
    echo "You will now be prompted for the STL OUTPUT folder"
    pause

    [void]$FolderBrowser.ShowDialog()

    $OutputFolder = $FolderBrowser.SelectedPath
}

##################

if(Test-Path -Path "C:\Program Files\OpenSCAD\openscad.exe"){
    $fileExe = "C:\Program Files\OpenSCAD\openscad.exe"
} else {
    cls
    echo "OpenSCAD .exe file not found in default location, you will need to select the .exe file manually"
    pause

    $FileBrowser.filter = "exe (*.exe)| *.exe"
    [void]$FileBrowser.ShowDialog()

    $fileExe = $FileBrowser.FileName
}

##################

$JsonPath = (Get-ItemProperty $scadPath).DirectoryName + '\' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".json"


##### START STL EXPORT #####
# Import JSON
$importedJSON = Get-Content $jsonPath | convertfrom-json

# Count Items
$totalItemCount = (Get-Member -MemberType NoteProperty -InputObject $importedJSON[0].parameterSets).count

# For Each Item
$importedJSON[0].parameterSets | Get-Member -MemberType NoteProperty | ForEach-Object {
    
    # Increase the count by 1
    $count += 1

    # Set the current file name to Parameter Set Name + .stl
    $Output_Filename = $_.Name + '.stl'
    # Set the Output Path
    $OutputPath = $OutputFolder + '\' + $Output_Filename

    ## Write Host
    write-host ("Exporting " + $Output_Filename)
    write-host ("Export Path " + $OutputPath)
    write-host "$count of $totalItemCount"

    ## Check if item exists, if not
    if (-not(Test-Path -Path $OutputPath)) {
        Write-Host "Starting..." 

        # Set the arguments
        $arguments = '-o "' + $OutputPath + '" -p "' + $jsonPath + '" -P "' + $_.Name + '" "' + $scadPath + '"'

        # Run the command
        start-process $fileExe $arguments -Wait

        Write-Host "Done..." 

    } else {
        # If already exists
        Write-Host "Already Exists..." 

    }

    Write-Host ""
    Write-Host "----"
    Write-Host ""

}

Write-Host "Finished"
pause