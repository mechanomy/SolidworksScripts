Attribute VB_Name = "exportDimensions"
'Exports every dimension in the active part or assembly to modelName.csv, in the same format as importProperties.
' Each dimension is named Part_Sketch_Name, eg bracket_Sketch1_D1, and its value is in document units.
' The active model is exported in all of its configurations; assembly components only in the configuration the assembly uses.
' Lightweight and suppressed components have no model loaded, so they are skipped.

'MIT License
'Copyright (c) 2026 Mechanomy
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim swApp As Object
Dim fOut As Object
Dim written As Object ' "config|Dimension.FullName" keys already written; a part used several times in an assembly is only exported once per configuration
Dim nWritten As Long

Sub main()
  Dim swModel As SldWorks.ModelDoc2
  Dim swAssy As SldWorks.AssemblyDoc
  Dim swComp As SldWorks.Component2
  Dim swCompModel As SldWorks.ModelDoc2
  Dim vComps As Variant
  Dim compConfig(0) As String
  Dim modelType As Long
  Dim pathModel As String
  Dim pathCsv As String
  Dim nSkipped As Long
  Dim i As Long

  Set swApp = Application.SldWorks
  Set swModel = swApp.ActiveDoc
  If swModel Is Nothing Then
    MsgBox "Open a part or assembly first, exiting", vbExclamation, "ExportDimensions"
    Exit Sub
  End If
  modelType = swModel.GetType
  If modelType <> swDocumentTypes_e.swDocPART And modelType <> swDocumentTypes_e.swDocASSEMBLY Then
    MsgBox "The active document must be a part or assembly, exiting", vbExclamation, "ExportDimensions"
    Exit Sub
  End If

  ' The CSV goes next to the model file, so the model must be saved first
  pathModel = swModel.GetPathName
  If pathModel = "" Then
    MsgBox "Save the model before running exportDimensions, exiting", vbExclamation, "ExportDimensions"
    Exit Sub
  End If
  pathCsv = Left(pathModel, InStrRev(pathModel, ".") - 1) & ".csv"
  Debug.Print "File = " & pathModel

  Set written = CreateObject("Scripting.Dictionary")
  nWritten = 0
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set fOut = fso.CreateTextFile(pathCsv, True)
  fOut.WriteLine "In [" & lengthUnitName(swModel.LengthUnit) & "] Properties written by " & swApp.GetCurrentMacroPathName ' importProperties skips this first line

  exportModel swModel, swModel.GetConfigurationNames

  If modelType = swDocumentTypes_e.swDocASSEMBLY Then
    Set swAssy = swModel
    vComps = swAssy.GetComponents(False) ' False = components at every level, including those inside subassemblies
    If Not IsEmpty(vComps) Then
      For i = 0 To UBound(vComps)
        Set swComp = vComps(i)
        Set swCompModel = swComp.GetModelDoc2 ' Nothing for lightweight and suppressed components
        If swCompModel Is Nothing Then
          Debug.Print "Skipped " & swComp.Name2 & ", it is lightweight or suppressed"
          nSkipped = nSkipped + 1
        Else
          compConfig(0) = swComp.ReferencedConfiguration
          exportModel swCompModel, compConfig
        End If
      Next i
    End If
  End If

  fOut.Close
  Set fOut = Nothing
  Set fso = Nothing
  Set written = Nothing

  Dim msg As String
  msg = "exportDimensions wrote " & nWritten & " dimensions to" & vbCrLf & pathCsv
  If nSkipped > 0 Then
    msg = msg & vbCrLf & vbCrLf & nSkipped & " lightweight or suppressed components were skipped, resolve them to include their dimensions"
  End If
  MsgBox msg & vbCrLf & vbCrLf & "Thank you for using Mechanomy", vbInformation, "ExportDimensions"
End Sub

' Writes the dimensions of every feature in swModel, and of the sketches absorbed into those features, for each of configNames
Sub exportModel(swModel As SldWorks.ModelDoc2, configNames As Variant)
  Dim swFeat As SldWorks.Feature
  Dim swSubFeat As SldWorks.Feature
  Dim modelName As String
  Dim pathModel As String

  ' GetTitle only includes the extension if Windows shows extensions, so take the name from the path
  pathModel = swModel.GetPathName
  modelName = Mid(pathModel, InStrRev(pathModel, "\") + 1)
  modelName = Left(modelName, InStrRev(modelName, ".") - 1)

  Set swFeat = swModel.FirstFeature
  Do While Not swFeat Is Nothing
    exportFeature swFeat, modelName, configNames
    Set swSubFeat = swFeat.GetFirstSubFeature ' eg the sketch under an extrude
    Do While Not swSubFeat Is Nothing
      exportFeature swSubFeat, modelName, configNames
      Set swSubFeat = swSubFeat.GetNextSubFeature
    Loop
    Set swFeat = swFeat.GetNextFeature
  Loop
End Sub

Sub exportFeature(swFeat As SldWorks.Feature, modelName As String, configNames As Variant)
  Dim swDispDim As SldWorks.DisplayDimension
  Dim swDim As SldWorks.Dimension
  Dim vValues As Variant
  Dim fullName As String
  Dim dimName As String
  Dim key As String
  Dim i As Long

  Set swDispDim = swFeat.GetFirstDisplayDimension
  Do While Not swDispDim Is Nothing
    Set swDim = swDispDim.GetDimension2(0)
    fullName = swDim.FullName ' eg D1@Sketch1@bracket.Part, the middle field is the feature that owns the dimension
    dimName = modelName & "_" & Split(fullName, "@")(1) & "_" & swDim.Name

    ' Values are in document units, one per requested configuration https://help.solidworks.com/2022/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IDimension~GetValue3.html
    vValues = swDim.GetValue3(swInConfigurationOpts_e.swSpecifyConfiguration, configNames)
    For i = 0 To UBound(configNames)
      key = configNames(i) & "|" & fullName
      If Not written.Exists(key) Then
        written.Add key, True
        fOut.WriteLine configNames(i) & "; " & dimName & "; double; " & vValues(i) & ";"
        nWritten = nWritten + 1
      End If
    Next i
    Set swDispDim = swFeat.GetNextDisplayDimension(swDispDim)
  Loop
End Sub

' Names the swLengthUnit_e value of a model's length unit https://help.solidworks.com/2022/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swLengthUnit_e.html
Function lengthUnitName(lengthUnit As Long) As String
  Select Case lengthUnit
  Case swLengthUnit_e.swMM
    lengthUnitName = "mm"
  Case swLengthUnit_e.swCM
    lengthUnitName = "cm"
  Case swLengthUnit_e.swMETER
    lengthUnitName = "m"
  Case swLengthUnit_e.swINCHES
    lengthUnitName = "in"
  Case swLengthUnit_e.swFEET
    lengthUnitName = "ft"
  Case swLengthUnit_e.swFEETINCHES
    lengthUnitName = "ft-in"
  Case swLengthUnit_e.swANGSTROM
    lengthUnitName = "angstrom"
  Case swLengthUnit_e.swNANOMETER
    lengthUnitName = "nm"
  Case swLengthUnit_e.swMICRON
    lengthUnitName = "um"
  Case swLengthUnit_e.swMIL
    lengthUnitName = "mil"
  Case swLengthUnit_e.swUIN
    lengthUnitName = "uin"
  Case Else
    lengthUnitName = "unit " & lengthUnit
  End Select
End Function
