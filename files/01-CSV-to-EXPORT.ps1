### You will be prompted for these values if not set here
$csvPath = "" ### Path to the CSV file
$scadPath = "" ### Path to the .SCAD file
$OutputFolder = "" ### Path of output folder
$file_extension = "" ### "STL", "OFF", "AMF", "3MF", "DXF", "SVG", "PNG"
$cam_args = "" ### Only needed if using PNG export
$process_count = 3

################ START OF SCRIPT
function pause{ $null = Read-Host 'Press Enter to continue...' }

if($IsWindows -or ((Get-Host | Select-Object Version).Version.Major -lt 7) ){

    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    
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

    If ($OutputFolder -eq ''){

        Clear-Host
        Write-Output "You will now be prompted for the STL OUTPUT folder"
        pause

        [void]$FolderBrowser.ShowDialog()

        $OutputFolder = $FolderBrowser.SelectedPath
    } elseif ($OutputFolder -notmatch '\\$'){
        $OutputFolder += '\'
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

        $file_extension = $file_extension_list | Sort-Object | Out-GridView @GridArguments | ForEach-Object {
            $_
        }

    }

    ##################

    if(Test-Path -Path "C:\Program Files\OpenSCAD\openscad.exe"){
        $fileExe = "C:\Program Files\OpenSCAD\openscad.exe"
    } else {
        Clear-Host
        Write-Output "OpenSCAD .exe file not found in default location, you will need to select the .exe file manually"
        pause

        $FileBrowser.filter = "exe (*.exe)| *.exe"
        [void]$FileBrowser.ShowDialog()

        $fileExe = $FileBrowser.FileName
    }

    ##################

    $JsonPath = (Get-ItemProperty $scadPath).DirectoryName + '\' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".json"

} elseif($IsLinux) {

    If ($csvPath -eq ''){
        Clear-Host
        $csvPath = Read-Host -Prompt 'Input the full CSV path'
    
    }
    
    ##################
    
    If ($scadPath -eq ''){
        Clear-Host
        $scadPath = Read-Host -Prompt 'Input the full .SCAD file path'
    }
    
    ##################
    
    If ($OutputFolder -eq ''){
    
        Clear-Host
        $OutputFolder = Read-Host -Prompt 'Input the export folder path'

        if ($OutputFolder -notmatch '\\$'){
            $OutputFolder += '\'
        }
    
        If(!(test-path $OutputFolder))
        {
            New-Item -ItemType Directory -Force -Path $OutputFolder
        }
   
    }
    
    ################
    
    If ($file_extension -eq ''){
    
        Clear-Host
        Write-Host "Enter one of the following options: STL, OFF, AMF, 3MF, DXF, SVG, PNG"
    
        $file_extension = Read-Host -Prompt 'Input the file extension'

    }

    ################

    $JsonPath = (Get-ItemProperty $scadPath).DirectoryName + '/' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".json"
    $fileExe = ' '
    
} else {
    Pause
    Write-Host "Max OS not yet supported"    
}

If ($file_extension -eq "PNG" -and $cam_args -eq ''){

Clear-Host
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

##################

if(Test-Path -Path $JsonPath){
    Clear-Host
    # Write-Output ".JSON file already exists, creating backup..."
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

# Write-Host "Done..."
# Write-Host "Starting Export..."




####### FILE EXPORT

if((Get-Host | Select-Object Version).Version.Major -gt 6){

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
            if($IsWindows){
                $OutputPath = $OutputFolder + '\' + $Output_Filename
            } else {
                $OutputPath = $OutputFolder + $Output_Filename
            }
        }

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
        $arguments = $PSItem.Value
        if($IsWindows){
            start-process $using:fileExe $arguments -Wait -RedirectStandardError NUL
        } else {
            start-process openscad $arguments -Wait -RedirectStandardError NUL
        }

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

} else {

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
            #write-host ("Export Path " + $OutputPath)

            Write-Progress -Activity ("Exporting " + $Output_Filename) -Status ("Progress: " + $current_count + " of " + $totalItemCount) -PercentComplete ($current_count/$totalItemCount*100)

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

        Clear-Host
        Write-Host "Exported: $export_count
    Skipped: $skip_count
    Total: $totalItemCount"

        pause

    } else {
        
        Clear-Host
        Write-host "Missing JSON file..."
        pause

    }

}
