Attribute VB_Name = "MateFTRModule"
'Mates the Front, Top, and Right planes of the selected component to the Front, Top, and Right planes of the active assembly
' Select the component in the feature tree, then run the macro
' Uses AddMate5 https://help.solidworks.com/2022/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IAssemblyDoc~AddMate5.html

'MIT License
'Copyright (c) 2026 Mechanomy
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim swApp As Object

Dim swAssy As Object
Dim swComp As Object
Dim swMate As Object
Dim boolstatus As Boolean
Dim longstatus As Long

Sub main()

    Set swApp = Application.SldWorks

    Set swAssy = swApp.ActiveDoc
    If swAssy Is Nothing Then Exit Sub
    If swAssy.GetType <> swDocumentTypes_e.swDocASSEMBLY Then
        MsgBox "MateFTR must be run from an assembly."
        Exit Sub
    End If

    ' Get the component that owns the current selection
    Set swComp = swAssy.SelectionManager.GetSelectedObjectsComponent4(1, -1)
    If swComp Is Nothing Then
        MsgBox "Select a component in the feature tree, then run MateFTR."
        Exit Sub
    End If

    ' The first three planes in any part or assembly are Front, Top, Right, even if renamed
    Dim assyPlanes As Variant
    Dim compPlanes As Variant
    assyPlanes = firstThreePlanes(swAssy.FirstFeature)
    compPlanes = firstThreePlanes(swComp.FirstFeature)

    Dim i As Integer
    For i = 0 To 2
        swAssy.ClearSelection2 True
        boolstatus = compPlanes(i).Select2(False, 1)
        boolstatus = assyPlanes(i).Select2(True, 1)
        Set swMate = swAssy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
        If swMate Is Nothing Then
            MsgBox "Failed to mate " & compPlanes(i).Name & " to " & assyPlanes(i).Name & ", error " & longstatus
        End If
    Next i

    swAssy.ClearSelection2 True
    swAssy.EditRebuild3

End Sub

' Returns the first three reference planes found by walking the feature tree from swFeat
Function firstThreePlanes(ByVal swFeat As Object) As Variant
    Dim planes(2) As Object
    Dim n As Integer
    n = 0
    Do While Not swFeat Is Nothing And n < 3
        If swFeat.GetTypeName2 = "RefPlane" Then
            Set planes(n) = swFeat
            n = n + 1
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    firstThreePlanes = planes
End Function
