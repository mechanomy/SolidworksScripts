# Macros for frequent Solidworks operations
Clone and add this repository to your [Macro File Locations](https://help.solidworks.com/2020/english/SolidWorks/sldworks/HIDD_Options_External_Folders.htm?format=P&value=) and mapping them to [shortcut keys](https://help.solidworks.com/2019/english/SolidWorks/sldworks/t_assigning_macro_keyboard_shortcut.htm) or [toolbar macro buttons](https://help.solidworks.com/2022/english/Solidworks/sldworks/t_assigning_macro_toolbar_button.htm).

Macros are written in Visual Basic (*.bas) and then complied by the VB editor.
To modify from Solidworks, Tools > Macro > New; tempName.swp; then in tree: Modules > right click on tempName > remove > no don't save; then right click on the project name > import file > select the .bas you want to edit.
Once edited, running the .bas will produce the .swp needed for shortcuts and buttons.
Clear as mud.

Toolbar button images for [macro buttons](https://help.solidworks.com/2022/english/Solidworks/sldworks/t_assigning_macro_toolbar_button.htm) must be 16x16 pixel BMP files with at most 256 colors and a white background ([ref](https://www.javelin-tech.com/blog/2020/10/creating-macro-buttons-in-solidworks/)).
Stock images are in the SOLIDWORKS install directory under `data\user macro icons`.

## addAxesXYZ
Adds canonical X,Y,Z axes to the open part/assembly.
This can also be done in the part/assembly template.

## export3mfStep
Saves the current part in 3MF and STEP files.

## importProperties2Part
Imports properties in a CSV file into the current part or assembly, as either 'Custom File Properties' or 'Configuration Properties'.
The CSV format is:
```csv
Optional reference path to file.csv or other comment; this first line is skipped when importing
Default; secretText; text;  configPropText;
Default; bigLength; double; 60.00000;
specialConfig; configNum; double; 2.500000;
specialConfig; configText; text;  specialest config;
```

## exportPartProperties
Exports the current parts properties to partName.csv in the same format as importProperties.

## exportDimensions
Exports every dimension in the current part or assembly to partName.csv in the same format as importProperties.
Each dimension is named `Part_Sketch_Name` (eg `bracket_Sketch1_D1`), typed `double`, and valued in document units.
The first line gives the model's length unit and the macro that wrote the file:
```csv
In [mm] Properties written by C:\SolidworksScripts\exportDimensions\exportDimensions.swp
Default; bracket_Sketch1_D1; double; 25;
Default; bracket_Boss-Extrude1_D1; double; 10;
```
The current model is exported in all of its configurations.
In an assembly, each component's dimensions are exported once, in the configuration the assembly uses, so the first column holds the component's configuration name.
Lightweight and suppressed components are skipped; resolve them to include their dimensions.

## saveIsometricPng
Saves an isometric, zoom-to-fit PNG of the current part next to the part file, with the same name. The previous view is restored afterwards.
The image is rendered as a BMP at `supersample` times the viewport size, then downscaled to the viewport size and saved as PNG with PowerShell, which anti-aliases the edges.
SOLIDWORKS' own PNG export is avoided because it writes alpha = 0 on every pixel, so the images display as blank.

## MateFTR
In an assembly, mates the Front, Top, and Right planes of each component selected in the feature tree to the assembly's Front, Top, and Right planes.
Planes are only mated by name (Front-Front, Top-Top, Right-Right); a component missing one of these planes is skipped for that plane.
Fixed components are floated before mating.

## Copyright
Copyright (c) 2026 Mechanomy LLC

