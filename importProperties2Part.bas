Attribute VB_Name = "importProperties2PartModule"
'This macro imports custom properties from a CSV file into the active part

'MIT License
'Copyright (c) 2023 Mechanomy
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim swApp As Object
Sub main()
    Debug.Print vbCrLf & vbCrLf & "--importProperties started at " & Now & "--"
    
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim i As Long
    Dim ret As Integer
    
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Debug.Print "File = " + swModel.GetPathName 'Debug.Prints appear in the VBA Immediate window
    
    Dim modelType As Integer
    Dim pathModel As String
    Dim pathCsv As String
    modelType = swModel.GetType
    pathModel = swModel.GetPathName
    If Len(pathModel) > 0 Then
        pathCsv = Left(pathModel, InStrRev(pathModel, ".") - 1) + ".csv" ' Strip the extension
    Else
        pathCsv = ""
    End If
    'Debug.Print "csv path: " + pathCsv
    
    ' https://help.solidworks.com/2016/english/api/sldworksapi/open_file_example_vb.htm
    Dim Filter As String
    Filter = "CSV (*.csv)|*.csv|All Files (*.*)|*.*|" 'format is [display text]|[filter] ...
    Dim fileName As String
    Dim fileConfig As String
    Dim fileDispName As String
    Dim fileOptions As Long
    
    'value = instance.GetOpenFileName(DialogTitle, InitialFileName, FileFilter, OpenOptions, ConfigName, DisplayName)
    'pathCsv = swApp.GetOpenFileName("Select properties file to import", pathCsv, Filter, fileOptions, fileConfig, fileDispName)
    pathCsv = swApp.GetOpenFileName("Select properties file to import", ".", Filter, fileOptions, fileConfig, fileDispName)
    Debug.Print pathCsv
    
    If modelType = swDocumentTypes_e.swDocPART And Len(pathCsv) > 0 Then
        'Open the CSV for import
        Dim fso As Object 'apparently this is the old, incorrect way...but it works  https://stackoverflow.com/questions/11503174/how-to-create-and-write-to-a-txt-file-using-vba
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim fCsv As Object
        Set fCsv = fso.OpenTextFile(pathCsv)
        
        strLine = fCsv.ReadLine 'the first line is the filepath, but we already know this by requiring the csv and sldprt to have the same names
        Dim nLine As Integer
        nLine = 2
    
        Do Until fCsv.AtEndOfStream
            strLine = fCsv.ReadLine
            If Len(strLine) > 5 Then 'ignore newlines
                'Debug.Print "strLine= " & strLine
            
                Dim splits() As String
                splits = Split(strLine, ";") 'expect format: Default; configPropNum; double; 1.300000
                
                Dim ii As Long
                If UBound(splits) + 1 = 5 Then 'Require the line to only have 4 elements
                    'Parse the line
                    Dim strConfig As String
                    strConfig = splits(0)
                    
                    Dim strName As String
                    strName = Trim(splits(1))
                    
                    Dim strType As String
                    Dim vPropType As Variant
                    strType = Trim(splits(2))
                    
                    Dim value As Variant
                    
                    Select Case strType 'type can be one of these enums: https://help.solidworks.com/2022/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swCustomInfoType_e.html
                    Case "date"
                        vPropType = swCustomInfoDate
                        value = splits(3) 'do we need type-specific parsing?
                    Case "double"
                        vPropType = swCustomInfoDouble
                        value = Trim(splits(3))
                    Case "integer"
                        vPropType = swCustomInfoNumber
                        value = Trim(splits(3))
                    Case "text"
                        vPropType = swCustomInfoText
                        value = splits(3)
                    Case "unknown"
                        vPropType = swCustomInfoUnknown
                        value = splits(3)
                    Case "yesOrNo"
                        vPropType = swCustomInfoYesOrNo
                        value = splits(3)
                    End Select
                    'Debug.Print "strType[" & strType & "] = vPropType[" & vPropType; "]"
                    
                    If Len(strConfig) = 0 Or strConfig = "Default" Then 'This first branch adds Custom [File] Properties
                        '241009 rewriting to use CustomPropertyManager:
                        Dim swModelDocExt As ModelDocExtension
                        Set swModelDocExt = swModel.Extension
                        Dim swCustProp As CustomPropertyManager ' list of members: https://help.solidworks.com/2021/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ICustomPropertyManager_members.html?_gl=1*q8mdhx*_up*MQ..*_ga*MTc5MTI0NzYuMTcyODUyNDU4OA..*_ga_XQJPQWHZHH*MTcyODUyNDU4OC4xLjEuMTcyODUyNTQxNy4wLjAuMA..
                        Set swCustProp = swModelDocExt.CustomPropertyManager("") 'Get the custom property data https://help.solidworks.com/2021/english/api/sldworksapi/Get_Custom_Properties_of_Referenced_Part_Example_VB.htm?_gl=1*1p8akwv*_up*MQ..*_ga*MTc5MTI0NzYuMTcyODUyNDU4OA..*_ga_XQJPQWHZHH*MTcyODUyNDU4OC4xLjEuMTcyODUyNDg4NC4wLjAuMA..
                        bool = swCustProp.Add3(strName, vPropType, "" & value, swCustomPropertyDeleteAndAdd)
                        
                    Else 'This branch adds configuration-specific properties

                        'Does the config already exist?
                        Dim configExists As Boolean
                        configExists = False
                        Dim vConfigs As Variant
                        vConfigs = swModel.GetConfigurationNames
                        For i = 0 To UBound(vConfigs)
                            Debug.Print "config " & i & " = " & vConfigs(i)
                            If Not configExists And strConfig = vConfigs(i) Then
                                configExists = True
                            End If
                        Next i
                        
                        Dim swConfig As SldWorks.Configuration
                        If configExists Then 'IsNull(swConfig) returns true even for nonexistant configs...
                            'Look up the config
                            'Debug.Print "Configuration [" & strConfig & "] exists"
                            Set swConfig = swModel.GetConfigurationByName(strConfig)
                        Else
                            Debug.Print "Configuration [" & strConfig & "] does not exist, creating"
                            Dim comment As String
                            Dim alt As String
                            Dim ops As swConfigurationOptions2_e
                            ops = swConfigOption_DontActivate 'add(+) desired options ... https://help.solidworks.com/2022/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swConfigurationOptions2_e.html
                            'value = instance.AddConfiguration3(Name, Comment, AlternateName, Options)
                            'ret = swModel.AddConfiguration3(strConfig, comment, alt, ops) 'https://help.solidworks.com/2022/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDoc2~AddConfiguration3.html
                            Set swConfig = swModel.AddConfiguration3(strConfig, "", "", swConfigOption_DontActivate)
                        End If
                        
                        Dim pOverwriteExisting As Integer
                        pOverwriteExisting = swCustomPropertyDeleteAndAdd 'https://help.solidworks.com/2022/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swCustomPropertyAddOption_e.html
                        
                        'These are configuration properties: https://help.solidworks.com/2022/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDocExtension~CustomPropertyManager.html
                        '...To access a general custom information value, set the configuration argument to an empty string. To get a document-level property, pass an empty string ("") to the configuration argument. (This doesn't work, hence the obsolete AddCustomInfo2() above.)
                        ret = swConfig.CustomPropertyManager.Add3(strName, vPropType, value, pOverwriteExisting) ' https://help.solidworks.com/2022/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.icustompropertymanager~add3.html , https://help.solidworks.com/2022/english/api/sldworksapi/Get_Custom_Properties_for_Configuration_Example_VB.htm
                        If ret = 0 Then '0=success, swCustomInfoAddResult_e https://help.solidworks.com/2022/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swCustomInfoAddResult_e.html
                            Debug.Print "Added [" & strName & "] = [" & value & "] to configuration[" & strConfig & "]"
                        Else
                            Debug.Print "Failed adding [" & strName & "] = [" & value & "] to configuration[" & strConfig & "]"
                        End If
                    End If
                Else
                    Debug.Print "Line [" & strLine & "] has incorrect delimiters, should be 'configuration; name; type; value;'"
                    MsgBox "Line " & nLine & ": [" & strLine & "] has incorrect delimiters, should be 'configuration; name; type; value;'", vbExclamation, "ImportProperties"
                End If 'splits length
            End If 'min line length
            nLine = nLine + 1
        Loop
        fCsv.Close
        MsgBox "importProperties finished " & vbCrLf & vbCrLf & "Thank you for using Mechanomy", vbInformation, "ImportProperties" ' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-function
    Else
        MsgBox "No file selected, exiting", vbExclamation, "ImportProperties"
    End If 'modelType=part
    
    Set fso = Nothing
    Set fCsv = Nothing
    Debug.Print "--importProperties finished at " & Now & "--"
End Sub

