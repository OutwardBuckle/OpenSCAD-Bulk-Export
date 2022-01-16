### You will be prompted for these values if not set here
$scadPath = "" ### Path to the .SCAD file
$OutputFolder = "" ### Path of output folder
$file_extension = "" ### "STL", "OFF", "AMF", "3MF", "DXF", "SVG", "PNG"
$cam_args = "" ### Only needed if using PNG export
$process_count = 3

################ START OF SCRIPT

function pause{ $null = Read-Host 'Press Enter to continue...' }

if($IsWindows){

    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog

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

$JsonPath = (Get-ItemProperty $scadPath).DirectoryName + '\' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".json"

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
