# Bulk export via OpenSCAD command line & PowerShell (Windows & Linux)

- [Script Functions](#script-functions)
- [CSV Format](#csv-format)
- [JSON Format](#json-format)
- [Command Line Parameters](#command-line-parameters)
- [How to run PowerShell scripts](#how-to-run-powershell-scripts)

## Script Functions

| Name          | Description                                                                            |
| --------------|:--------------------------------------------------------------------------------------:|
| CSV to EXPORT | Import a CSV file and export each item (to STL, OFF, AMF, 3MF, DXF, SVG or PNG)        |
| CSV to JSON   | Create a customizer parameter set JSON file from a CSV file                            |
| JSON to EXPORT| Export all parameter sets from a JSON file                                             |
| SCAD to CSV   | Generate a CSV from a SCAD file (currently requires the latest dev build of OpenSCAD)  |

## CSV Format

The CSV file only requires one field (exported_filename), all other fields should relate to values in the .SCAD script. Any varibles that are not specified in the CSV will use the default values from the SCAD file.

_Example .SCAD File:_

    message = "Some Text";
    myfont = "Stencil";
    textsize = 20;
    height = 10;

    linear_extrude(height){
        text(message, font=myfont, size=textsize);
    }

_Example .CSV File_

| exported_filename | message   | myfont  |
| ------------------|:---------:| -------:|
| Hi-cali           | Hi        | Calibri |
| Hello-cali        | Hello     | Calibri |
| Howdy-Stencil     | Howdy     | Stencil |

_Example Output_

![image](https://user-images.githubusercontent.com/50000826/140439376-16148446-163c-4ac3-9986-237b54ac9945.png)

## JSON Format

The script uses the OpenSCAD customizer function to run the exports, so the JSON file matches the format of this - https://en.wikibooks.org/wiki/OpenSCAD_User_Manual/Customizer#Saving_Parameters_value_in_JSON_file

This allows you to create/save many different configurations using the OpenSCAD customizer GUI, then easily export them all to your desired file format.

_If you're exporting from CSV and you already have a JSON file with the same name as the chosen .SCAD file, it will be renamed with a timestamp at the start (OpenSCAD needs the JSON file to have the same filename as the .SCAD file, or it will export the files with default parameters)_

More info on parameter sets, etc: https://en.wikibooks.org/wiki/OpenSCAD_User_Manual/Customizer

## Command Line Parameters

If you run the script without any parameters, you'll be prompted to enter the values, otherwise you can set them when calling the script:

* __scadPath__ - Path to the .SCAD file
* __inputType__ - Select an input type. Valid options are: _CSV_, _JSON_ or _SCAD_
* __fileExtension__ - Extension of the output files. Valid options are: _STL_, _OFF_, _AMF_, _3MF_, _DXF_, _SVG_, _PNG_, _CSV_ or _JSON_
* __csvPath__ - Path to the .SCAD file (Only required if exporting to/from .CSV)
* __outputFolder__ - Path to export the files (Must end with a forward slash on linux. Not required if exporting to .JSON or .CSV)
* __camArgs__ - Camera arguments for PNG export (Only required if exporting to .PNG)
* __overwriteFiles__ - Existing files will not be overwritten unless set to _true_. Valid options are: _$True_ or _$False_
* __processCount__ - Number of exports to run at a time (Defaults to 3. Must be using PowerShell V7 or later)

### Examples:

CSV TO JSON

    .\OpenSCAD_Bulk_Exporter.ps1 -scadPath Example.scad -inputType CSV -csvPath Example_CSV.csv -fileExtension JSON

JSON TO EXPORT

    .\OpenSCAD_Bulk_Exporter.ps1 -scadPath Example.scad -inputType JSON -outputFolder 'Output' -fileExtension STL -overwriteFiles $True

CSV TO JSON TO EXPORT

    .\OpenSCAD_Bulk_Exporter.ps1 -scadPath Example.scad -inputType CSV -csvPath Example_CSV.csv -outputFolder 'Output' -fileExtension STL

## How to run PowerShell scripts

### On Windows:
* Right-click the file
* Select Run With PowerShell

Or

Run it from command line:

    .\OpenSCAD_Bulk_Exporter.ps1 -scadPath Example.scad -inputType CSV -csvPath Example_CSV.csv -outputFolder 'Output' -fileExtension STL

### On Linux:
* [Install PowerShell](https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-linux?view=powershell-7.2)

From Terminal:

    pwsh -file OpenSCAD_Bulk_Exporter.ps1

Or start a session by running ```pwsh``` and then enter a dot followed by the full script path in quotes, for example:

    ."/home/localuser/Downloads/OpenSCAD_Bulk_Exporter.ps1" -scadPath Example.scad -inputType CSV -csvPath Example_CSV.csv -outputFolder 'Output/' -fileExtension STL

### Note

PowerShell 5 is already installed on Windows 10, however if using linux or if you want to run multiple exports in parallel, you'll need to install PowerShell 7
* [Install on Windows](https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.1#msi)
* [Install on Linux](https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-linux?view=powershell-7.2)
