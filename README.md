# OpenSCAD Bulk Export PowerShell Script

Bulk export via OpenSCAD command line and PowerShell

I've included a very basic SCAD and CSV file so you can see how they work

To run, download the ps1 file, right click it and select Run With PowerShell. It will:

1. Prompt you for 
    - a .CSV file
    - a .SCAD file
    - The output folder for the STLs
2. Create a JSON file full of parameter sets for each line in the csv
3. Export each set as an STL file

The CSV file only requires one field (exported_filename), all other fields should relate to values in the .SCAD script. Any varibles that are not specified in the CSV will use the default values.

A quick note:
* OpenSCAD needs the JSON file to have the same filename as the .SCAD file, or it will export the STLs with default parameters (e.g. File_Name.SCAD & File_Name.JSON)
* If you already have a JSON file with the same name as the chosen .SCAD file, it will be renamed with a timestamp at the start

More info on parameter sets, etc: https://en.wikibooks.org/wiki/OpenSCAD_User_Manual/Customizer
