# Macros for frequent Solidworks operations
I suggest adding this repository to your [Macro File Locations](https://help.solidworks.com/2020/english/SolidWorks/sldworks/HIDD_Options_External_Folders.htm?format=P&value=) and mapping them to [shortcut keys](https://help.solidworks.com/2019/english/SolidWorks/sldworks/t_assigning_macro_keyboard_shortcut.htm).

As SWP is a binary format, the macros are exported as BAS files for clarity.

Toolbar button images for [macro buttons](https://help.solidworks.com/2022/english/Solidworks/sldworks/t_assigning_macro_toolbar_button.htm) must be 16x16 pixel BMP files with at most 256 colors and a white background ([ref](https://www.javelin-tech.com/blog/2020/10/creating-macro-buttons-in-solidworks/)); export3mfStep.bmp is a 16-color example.
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
Exports the current parts properties to partName.csv in the same format as importProperties2Part.

## savePNG
Saves an isometric, zoom-to-fit PNG of the current part next to the part file, with the same name. The previous view is restored afterwards.

## MateFTR
In an assembly, mates the Front, Top, and Right planes of each component selected in the feature tree to the assembly's Front, Top, and Right planes.
Planes are only mated by name (Front-Front, Top-Top, Right-Right); a component missing one of these planes is skipped for that plane.
Fixed components are floated before mating.

## Copyright
Copyright (c) 2023 Mechanomy LLC

