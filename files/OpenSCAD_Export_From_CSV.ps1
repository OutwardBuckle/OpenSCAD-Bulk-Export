Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
function pause{ $null = Read-Host 'Press Any Key or Enter to continue...' }

################

cls
echo "You will now be prompted for the CSV file"
pause

$FileBrowser.filter = "csv (*.csv)| *.csv"
[void]$FileBrowser.ShowDialog()

$csvPath = $FileBrowser.FileName

##################

cls
echo "You will now be prompted for the SCAD file"
pause

$FileBrowser.filter = "scad (*.scad)| *.scad"
[void]$FileBrowser.ShowDialog()

$scadPath = $FileBrowser.FileName

##################

cls

echo "You will now be prompted for the STL OUTPUT folder"
pause

[void]$FolderBrowser.ShowDialog()

$OutputFolder = $FolderBrowser.SelectedPath


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
$backupJsonPath = (Get-ItemProperty $scadPath).DirectoryName + '\' + (Get-Date -Format 'yyMMdd-hhmm') + '-' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".json"

if(Test-Path -Path $JsonPath){
    cls
    echo ".JSON file already exists, creating backup..."
    Copy-Item -Path $JsonPath -Destination ((Get-ItemProperty $scadPath).DirectoryName + '\' + (Get-Date -Format 'yyMMdd-hhmm') + '-' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".json")
    Write-Host "----"
    Write-Host ""
} 

##### START JSON EXPORT #####
echo "Creating JSON File..."
# Build the json structure
$jsonBase = @{}
$parameterSets = @{}

# Import the CSV
$csvData = Import-Csv -Path $csvPath

# For each item:
$csvData | foreach {
    
    # Set the name for the current parameter set
    $parameterSetName = $_.exported_filename
    
    # Add the current set to the array
    $parameterSets.Add($parameterSetName,$_)
}

# Add the items to the json base
$jsonBase.Add("parameterSets",$parameterSets)

# Output to a file
$jsonBase | ConvertTo-Json | Set-Content $jsonPath



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