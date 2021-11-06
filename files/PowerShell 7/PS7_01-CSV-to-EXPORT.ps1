### You will be prompted for these values if not set here
$csvPath = "D:\Storage\3D Printing\Thingiverse_Uploads\OpenSCAD Bulk Scripts\GitHub\files\Example_CSV.csv" ### Path to the CSV file
$scadPath = "D:\Storage\3D Printing\Thingiverse_Uploads\OpenSCAD Bulk Scripts\GitHub\files\Example.scad" ### Path to the .SCAD file
$OutputFolder = "D:\Storage\3D Printing\Thingiverse_Uploads\OpenSCAD Bulk Scripts\GitHub\files\PS7\output" ### Path of output folder
$file_extension = "" ### "STL", "OFF", "AMF", "3MF", "DXF", "SVG", "PNG"
$cam_args = "" ### Only needed if using PNG export
$process_count = 3

################ START OF SCRIPT

Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
function pause{ $null = Read-Host 'Press Enter to continue...' }

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
    # echo ".JSON file already exists, creating backup..."
    Copy-Item -Path $JsonPath -Destination ((Get-ItemProperty $scadPath).DirectoryName + '\' + (Get-Date -Format 'yyMMdd-hhmm') + '-' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".json")
    Write-Host "----"
    Write-Host ""
} 

##### START JSON EXPORT #####
# Write-Host "Creating JSON File..."
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

# Write-Host "Done..."
# Write-Host "Starting Export..."

####### FILE EXPORT
$importedJSON = Get-Content $jsonPath | convertfrom-json
$totalItemCount = (Get-Member -MemberType NoteProperty -InputObject $importedJSON[0].parameterSets).count
$dataset = @()

$importedJSON[0].parameterSets | Get-Member -MemberType NoteProperty | ForEach-Object {
    $index += 1
    # Set the current file name to Parameter Set Name + .stl
    $Output_Filename = $_.Name + '.' + $file_extension
    # Set the Output Path
    if ($OutputFolder -match '\\$'){
        $OutputPath = $OutputFolder + $Output_Filename
    } else {
            $OutputPath = $OutputFolder + '\' + $Output_Filename
    }

    ## Check if item exists, if not


    # Set the arguments
    If ($file_extension -eq "PNG"){
        $arguments = '-o "' + $OutputPath + '" ' + $cam_args + ' -p "' + $jsonPath + '" -P "' + $_.Name + '" "' + $scadPath + '"'
    } else {
        $arguments = '-o "' + $OutputPath + '" -p "' + $jsonPath + '" -P "' + $_.Name + '" "' + $scadPath + '"'
    }

    # Run the command
    $current_data_set = @{
        Id=$index
        Value=$arguments
        Name=$_.Name
    }
    $dataset += $current_data_set
}

# Create a hashtable for process.
# Keys should be ID's of the processes
$origin = @{}
$dataset | Foreach-Object {$origin.($_.id) = @{}}

# Create synced hashtable
$sync = [System.Collections.Hashtable]::Synchronized($origin)

$job = $dataset | Foreach-Object -ThrottleLimit $process_count -AsJob -Parallel {
    $syncCopy = $using:sync
    $process = $syncCopy.$($PSItem.Id)

    $process.Id = $PSItem.Id
    $process.Activity = "$($PSItem.Name).$using:file_Extension - $($PSItem.Id) of $using:totalItemCount"
    $process.Status = "Processing"

    #$PSItem.Value
    $args = $PSItem.Value
    start-process $using:fileExe $args -Wait -RedirectStandardError NUL

    # Mark process as completed
    $process.Completed = $true
}

# Show progress
while($job.State -eq 'Running')
{
    $sync.Keys | Foreach-Object {
        # If key is not defined, ignore
        if(![string]::IsNullOrEmpty($sync.$_.keys))
        {
            # Create parameter hashtable to splat
            $param = $sync.$_

            # Execute Write-Progress
            Write-Progress @param
        }
    }

    # Wait to refresh to not overload gui
    Start-Sleep -Seconds 0.1
}

# Write-Host "Done..."
pause