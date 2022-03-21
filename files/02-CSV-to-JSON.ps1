### You will be prompted for these values if not set here
$csvPath = ""
$scadPath = ""

################ START OF SCRIPT

function pause{ $null = Read-Host 'Press Enter to continue...' }

if($IsWindows -or ((Get-Host | Select-Object Version).Version.Major -lt 7) ){

    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog

    ################

    If ($csvPath -eq ''){

        Clear-Host
        Write-Output "You will now be prompted for the CSV file"
        pause

        $FileBrowser.filter = "csv (*.csv)| *.csv"
        [void]$FileBrowser.ShowDialog()

        $csvPath = $FileBrowser.FileName

    }

    ##################

    If ($scadPath -eq ''){

        Clear-Host
        Write-Output "You will now be prompted for the SCAD file"
        pause

        $FileBrowser.filter = "scad (*.scad)| *.scad"
        [void]$FileBrowser.ShowDialog()

        $scadPath = $FileBrowser.FileName
    }

    ##################

    $JsonPath = (Get-ItemProperty $scadPath).DirectoryName + '\' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".json"

} elseif($IsLinux) {

    If ($null -eq $csvPath){
        Clear-Host
        $csvPath = Read-Host -Prompt 'Input the full CSV path'
    
    }
    
    ##################
    
    If ($scadPath -eq ''){
        Clear-Host
        $scadPath = Read-Host -Prompt 'Input the full .SCAD file path'
    }

    ##################

    $JsonPath = (Get-ItemProperty $scadPath).DirectoryName + '/' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".json"

} else {
    Pause
    Write-Host "Max OS not yet supported"    
}

##################

if(Test-Path -Path $JsonPath){
    Clear-Host
    Write-Output ".JSON file already exists, creating backup..."
    Copy-Item -Path $JsonPath -Destination ((Get-ItemProperty $scadPath).DirectoryName + '\' + (Get-Date -Format 'yyMMdd-hhmm') + '-' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".json")
    Write-Host "----"
    Write-Host ""
} 

##### START JSON EXPORT #####
Write-Output "Creating JSON File..."
# Build the json structure
$jsonBase = @{}
$parameterSets = @{}

# Import the CSV
$csvData = Import-Csv -Path $csvPath

# For each item:
$csvData | ForEach-Object {
    
    # Set the name for the current parameter set
    $parameterSetName = $_.exported_filename
    
    # Add the current set to the array
    $parameterSets.Add($parameterSetName,$_)
}

# Add the items to the json base
$jsonBase.Add("parameterSets",$parameterSets)

# Output to a file
$jsonBase | ConvertTo-Json | Set-Content $jsonPath

Clear-Host
write-host "JSON file has been created:
$jsonPath"
pause
