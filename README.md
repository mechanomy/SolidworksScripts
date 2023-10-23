# Macros for frequent Solidworks operations
I suggest adding this repository to your [Macro File Locations](https://help.solidworks.com/2020/english/SolidWorks/sldworks/HIDD_Options_External_Folders.htm?format=P&value=) and mapping them to [shortcut keys](https://help.solidworks.com/2019/english/SolidWorks/sldworks/t_assigning_macro_keyboard_shortcut.htm).

As SWP is a binary format, the macros are exported as BAS files for clarity.

## addAxesXYZ
Adds canonical X,Y,Z axes to the open part/assembly.
This can also be done in the part/assembly template.

## export3mfStep
Saves the current part in 3MF and STEP files.

## importProperties2Part
Imports properties in a CSV file into the current part, as either 'Custom File Properties' or 'Configuration Properties'.
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

## Copyright
Copyright (c) 2023 Mechanomy LLC

