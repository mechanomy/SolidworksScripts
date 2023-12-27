Attribute VB_Name = "exportPartPropertiesModule"
'Exports model properties to filename.csv

'MIT License
'Copyright (c) 2023 Mechanomy
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim swApp As Object
Sub main()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Debug.Print "File = " + swModel.GetPathName 'Debug.Print s appear in the VBA Immediate window
    
    Dim modelType As Integer
    Dim pathModel As String
    Dim pathCsv As String
    modelType = swModel.GetType
    pathModel = swModel.GetPathName
    pathCsv = Left(pathModel, InStrRev(pathModel, ".") - 1) + ".csv" ' Strip the extension
    'Debug.Print "bare path: " + pathCsv
        
    Dim fso As Object 'apparently this is the old, incorrect way...but it works  https://stackoverflow.com/questions/11503174/how-to-create-and-write-to-a-txt-file-using-vba
    'Dim fso As New FileSystemObject 'not defined
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fOut As Object
    Set fOut = fso.CreateTextFile(pathCsv)
    fOut.WriteLine pathModel 'https://learn.microsoft.com/en-us/previous-versions/tn-archive/ee198716(v=technet.10)?redirectedfrom=MSDN
        
    Dim swConfig As SldWorks.Configuration
    Dim vConfName As Variant
    Dim vPropName As Variant
    Dim vPropValue As Variant
    Dim vPropType As Variant
    Dim nNumProp As Long
    Dim i As Long
    Dim j As Long
    vConfName = swModel.GetConfigurationNames
    If modelType = swDocumentTypes_e.swDocPART Then
        For i = 0 To UBound(vConfName)
            Set swConfig = swModel.GetConfigurationByName(vConfName(i))
            nNumProp = swConfig.GetCustomProperties(vPropName, vPropValue, vPropType) 'https://help.solidworks.com/2022/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.iconfiguration~getcustomproperties.html
            
            For j = 0 To nNumProp - 1
                Dim strType As String
                
                Debug.Print "j[" & j & "]:" & vPropType(j)
                
                Select Case vPropType(j)
                Case swCustomInfoDate
                    strType = "date"
                Case swCustomInfoDouble
                    strType = "double"
                Case swCustomInfoNumber
                    strType = "integer"
                Case swCustomInfoText
                    strType = "text"
                Case swCustomInfoUnknown
                    strType = "unknown"
                Case swCustomInfoYesOrNo
                    strType = "yesOrNo"
                End Select
            
                'Debug.Print vConfName(i) & ": " & vPropName(j) & " [" & vPropType(j) & " == " & strType & "] = " & vPropValue(j)
                fOut.WriteLine vConfName(i) & "; " & vPropName(j) & "; " & strType & "; " & vPropValue(j) & ";" 'use vbTab for tab separators
            Next j
        Next i
    End If
    fOut.Close
    Set fso = Nothing
    Set fOut = Nothing
    MsgBox "writeProperties wrote " & vbCrLf & pathCsv & vbCrLf & vbCrLf & "Thank you for using Mechanomy", vbInformation, "ExportProperties"

End Sub

