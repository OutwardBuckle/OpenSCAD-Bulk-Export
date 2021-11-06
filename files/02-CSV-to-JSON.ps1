### You will be prompted for these values if not set here
$csvPath = ""
$scadPath = ""

################ START OF SCRIPT

Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
function pause{ $null = Read-Host 'Press Any Key or Enter to continue...' }

################

If ($csvPath -eq ''){

    cls
    echo "You will now be prompted for the CSV file"
    pause

    $FileBrowser.filter = "csv (*.csv)| *.csv"
    [void]$FileBrowser.ShowDialog()

    $csvPath = $FileBrowser.FileName

}

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

$JsonPath = (Get-ItemProperty $scadPath).DirectoryName + '\' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".json"

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

cls
write-host "JSON file has been created:
$jsonPath"
pause
