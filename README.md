# Bulk export via OpenSCAD command line & PowerShell

File Descriptions
-----

| Name             | Description                                                                            |
| -----------------|:--------------------------------------------------------------------------------------:|
| 01-CSV-to-EXPORT | Import a CSV file and export each item (to STL, OFF, AMF, 3MF, DXF, SVG or PNG)        |
| 02-CSV-to-JSON   | Create  a customizer parameter set JSON file from a CSV file                           |
| 03-JSON-to-EXPORT| Export all a parameter sets from a JSON file                                           |
| 04-SCAD-to-CSV   | Generate a CSV from a SCAD file (currently requires the latest dev build of OpenSCAD)  |

Quick Instructions
-----

Download 01-CSV-to-EXPORT.ps1 from the files directory, right click it and select Run With PowerShell. It will:

1. Prompt you for 
    - a .CSV file
    - a .SCAD file
    - The output folder for the exported files (note: the prompt sometimes hides behind programs)
    - The exported file format
    - If you've selected PNG, it will also prompt for camera arguments
2. Create a JSON file full of parameter sets for each line in the csv
3. Export each item

The CSV file only requires one field (exported_filename), all other fields should relate to values in the .SCAD script. Any varibles that are not specified in the CSV will use the default values from the SCAD file.

CSV Format
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

![image](https://user-images.githubusercontent.com/50000826/140439376-16148446-163c-4ac3-9986-237b54ac9945.png)


-----

Notes
-----
* OpenSCAD needs the JSON file to have the same filename as the .SCAD file, or it will export the STLs with default parameters (e.g. File_Name.SCAD & File_Name.JSON)
* If you already have a JSON file with the same name as the chosen .SCAD file, it will be renamed with a timestamp at the start
* You can hard-code values at the top of each script file if you don't want to be prompted each time

More info on parameter sets, etc: https://en.wikibooks.org/wiki/OpenSCAD_User_Manual/Customizer
