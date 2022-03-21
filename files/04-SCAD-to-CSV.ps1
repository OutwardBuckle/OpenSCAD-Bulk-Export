### You will be prompted for these values if not set here
$scadPath = ""

################ START OF SCRIPT

function pause{ $null = Read-Host 'Press Any Key or Enter to continue...' }
if($IsWindows -or ((Get-Host | Select-Object Version).Version.Major -lt 7) ){

    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog

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

    $fileExe = "C:\Program Files\OpenSCAD Dev\openscad.exe" # Requires latest dev build as of 2021-11

    If ($fileExe -eq ''){

        Clear-Host
        Write-Output "Select the OpenSCAD .exe file (requires latest dev build as of 2021-11)"
        pause

        $FileBrowser.filter = "exe (*.exe)| *.exe"
        [void]$FileBrowser.ShowDialog()

        $fileExe = $FileBrowser.FileName

    }

    ##################

    $CSVfile = (Get-ItemProperty $scadPath).DirectoryName + '\' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".csv"

} elseif($IsLinux) {

    If ($scadPath -eq ''){
        Clear-Host
        $scadPath = Read-Host -Prompt 'Input the full .SCAD file path'
    }

    ##################

    $CSVfile = (Get-ItemProperty $scadPath).DirectoryName + '/' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".json"

} else {
    Pause
    Write-Host "Max OS not yet supported"    
}

##################


if(Test-Path -Path $CSVfile){
    $CSVfile = (Get-ItemProperty $scadPath).DirectoryName + '\' + (Get-Date -Format 'yyMMdd-hhmm') + '-' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".csv"
}


# Path for temporary json output file
$tempJsonPath = (Get-ItemProperty $scadPath).DirectoryName + '\' + (Get-Date -Format 'yyMMdd-hhmm') + '-temp-' + [System.IO.Path]::GetFileNameWithoutExtension($scadPath) + ".json"
# Set the arguments
$arguments = '--export-format param -o "' + $tempJsonPath + '" "' + $scadPath + '"'
# Run the command
if($IsWindows){
    start-process $using:fileExe $arguments -Wait
} else {
    start-process openscad $arguments -Wait
}
# Build table with name
$myObject = [PSCustomObject]@{
    exported_filename     = 'default'
}
# Add each item to the table
(Get-Content $tempJsonPath | ConvertFrom-Json).parameters | ForEach-Object {
    $myObject | Add-Member -MemberType NoteProperty -Name $_.name -Value $_.initial
}
# Export to CSV
$myObject | export-csv -Path $CSVfile -NoTypeInformation
# Cleanup JSON file
Remove-Item -Path $tempJsonPath

Clear-Host
write-host "CSV file has been created:
$CSVfile"

pause
