# Bulk export via OpenSCAD command line & PowerShell

Download OpenSCAD_Bulk_Export_0-AIO.ps1 file, right click it and select Run With PowerShell. It will:

1. Prompt you for 
    - a .CSV file
    - a .SCAD file
    - The output folder for the STLs
2. Create a JSON file full of parameter sets for each line in the csv
3. Export each set as an STL file

I've also included two extra files if you only want to go from CSV to JSON or from JSON to STL.

Update: Added another to generate a CSV from a SCAD file, currently requires the latest dev build of OpenSCAD.

The CSV file only requires one field (exported_filename), all other fields should relate to values in the .SCAD script. Any varibles that are not specified in the CSV will use the default values from the SCAD file.

-----

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

![](https://github.com/OutwardBuckle/OpenSCAD-Bulk-Export/blob/main/img/eg.png?raw=true)

-----

Quick notes:
* OpenSCAD needs the JSON file to have the same filename as the .SCAD file, or it will export the STLs with default parameters (e.g. File_Name.SCAD & File_Name.JSON)
* If you already have a JSON file with the same name as the chosen .SCAD file, it will be renamed with a timestamp at the start

More info on parameter sets, etc: https://en.wikibooks.org/wiki/OpenSCAD_User_Manual/Customizer
