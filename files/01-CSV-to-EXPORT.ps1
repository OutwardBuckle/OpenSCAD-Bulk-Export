### You will be prompted for these values if not set here
$csvPath = "" ### Path to the CSV file
$scadPath = "" ### Path to the .SCAD file
$OutputFolder = "" ### Path of output folder
$file_extension = "" ### "STL", "OFF", "AMF", "3MF", "DXF", "SVG", "PNG"
$cam_args = "" ### Only needed if using PNG export

################ START OF SCRIPT

Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
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

If ($OutputFolder -eq ''){

    cls
    echo "You will now be prompted for the STL OUTPUT folder"
    pause

    [void]$FolderBrowser.ShowDialog()

    $OutputFolder = $FolderBrowser.SelectedPath
}

If(!(test-path $OutputFolder))
{
      New-Item -ItemType Directory -Force -Path $OutputFolder
}

##################


If ($file_extension -eq ''){

    $file_extension_list = @(
        "STL", "OFF", "AMF", "3MF", "DXF", "SVG", "PNG"
    )

    $GridArguments = @{
        OutputMode = 'Single'
        Title      = 'Please select a export format and click OK'
    }

    $file_extension = $file_extension_list | Sort-Object | Out-GridView @GridArguments | foreach {
        $_
    }



    If ($file_extension -eq "PNG" -and $cam_args -eq ''){

    cls
    Write-Host "Default Camera Settings
    Leave Blank

Specific Camera Location ( translate_x,y,z, rot_x,y,z, dist)
    --camera 0,0,0,120,0,50,140

Auto-center
    --autocenter

View All
    --viewall
 
Color Scheme
    --colorscheme DeepOcean

User Render Quality
    --render

Set Image Export Size
    --imgsize 100,100



Example - Top Down, View All, Auto-Center, Deep Ocean Color Scheme
    --camera 0,0,0,0,0,0,0 --viewall --autocenter --colorscheme DeepOcean

Example - Default camera position, View All, Auto-Center, Deep Ocean Color Scheme
    --viewall --autocenter --imgsize 1024,1024 --render --colorscheme DeepOcean
    
"

    $cam_args = Read-Host -Prompt 'Input the camera arguments then press enter'

    }

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
If(Test-Path -Path $JsonPath){

    $importedJSON = Get-Content $jsonPath | convertfrom-json

    # Count Items
    $totalItemCount = (Get-Member -MemberType NoteProperty -InputObject $importedJSON[0].parameterSets).count

    # For Each Item
    $importedJSON[0].parameterSets | Get-Member -MemberType NoteProperty | ForEach-Object {
    
        # Increase the count by 1
        $current_count += 1

        # Set the current file name to Parameter Set Name + .stl
        $Output_Filename = $_.Name + '.' + $file_extension
        # Set the Output Path
        if ($OutputFolder -match '\\$'){
            $OutputPath = $OutputFolder + $Output_Filename
        } else {
             $OutputPath = $OutputFolder + '\' + $Output_Filename
        }

        ## Write Host
        cls
        write-host "
    
    



    
        "
        write-host ("Exporting " + $Output_Filename)
        write-host ("Export Path " + $OutputPath)
        write-host "$current_count of $totalItemCount"

        Write-Progress -Activity "Exporting" -Status "Progress:" -PercentComplete ($current_count/$totalItemCount*100)

        ## Check if item exists, if not
        if (-not(Test-Path -Path $OutputPath)) {

            # Set the arguments
            If ($file_extension -eq "PNG"){
                $arguments = '-o "' + $OutputPath + '" ' + $cam_args + ' -p "' + $jsonPath + '" -P "' + $_.Name + '" "' + $scadPath + '"'
            } else {
                $arguments = '-o "' + $OutputPath + '" -p "' + $jsonPath + '" -P "' + $_.Name + '" "' + $scadPath + '"'
            }

            # Run the command
            start-process $fileExe $arguments -Wait -RedirectStandardError NUL

            $export_count += 1

        } else {
             $skip_count += 1
        }

    }

    cls
    Write-Host "Exported: $export_count
Skipped: $skip_count
Total: $totalItemCount"

    pause

} else {
    
    cls
    Write-host "Missing JSON file..."
    pause

}