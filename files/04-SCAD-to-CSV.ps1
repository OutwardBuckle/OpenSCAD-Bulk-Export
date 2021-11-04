### You will be prompted for these values if not set here
$scadPath = ""
$fileExe = "" # Requires latest dev build as of 2021-11

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

If ($fileExe -eq ''){

    cls
    echo "Select the OpenSCAD .exe file (requires latest dev build as of 2021-11)"
    pause

    $FileBrowser.filter = "exe (*.exe)| *.exe"
    [void]$FileBrowser.ShowDialog()

    $fileExe = $FileBrowser.FileName

}

##################

# Output CSV File
$CSVfile = (Get-ItemProperty $scadPath).DirectoryName + '\' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".csv"
# Path for temporary json output file
$tempJsonPath = (Get-ItemProperty $scadPath).DirectoryName + '\' + (Get-Date -Format 'yyMMdd-hhmm') + '-temp-' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".json"
# Set the arguments
$arguments = '--export-format param -o "' + $tempJsonPath + '" "' + $scadPath + '"'
# Run the command
start-process $fileExe $arguments -Wait
# Build table with name
$myObject = [PSCustomObject]@{
    exported_filename     = 'Example'
}
# Add each item to the table
(Get-Content $tempJsonPath | ConvertFrom-Json).parameters | foreach {
    $myObject | Add-Member -MemberType NoteProperty -Name $_.name -Value $_.initial
}
# Export to CSV
$myObject | export-csv -Path $CSVfile -NoTypeInformation
# Cleanup JSON file
Remove-Item -Path $tempJsonPath