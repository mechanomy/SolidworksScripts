Attribute VB_Name = "MateFTRModule"
'Mates the Front, Top, and Right planes of each selected component to the Front, Top, and Right planes of the active assembly
' Select one or more components in the feature tree, then run the macro
' Uses AddMate5 https://help.solidworks.com/2022/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IAssemblyDoc~AddMate5.html

'MIT License
'Copyright (c) 2026 Mechanomy
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim swApp As Object

Dim swAssy As Object
Dim swSelMgr As Object
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

    ' Collect the selected components before mating clears the selection
    Set swSelMgr = swAssy.SelectionManager
    Dim comps As New Collection
    Dim i As Integer
    Dim j As Integer
    Dim isNew As Boolean
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        Set swComp = swSelMgr.GetSelectedObjectsComponent4(i, -1)
        If Not swComp Is Nothing Then
            isNew = True
            For j = 1 To comps.Count
                If comps(j) Is swComp Then isNew = False
            Next j
            If isNew Then comps.Add swComp
        End If
    Next i
    If comps.Count = 0 Then
        MsgBox "Select one or more components in the feature tree, then run MateFTR."
        Exit Sub
    End If

    ' Mate each plane only to the plane of the same name
    Dim planeNames As Variant
    planeNames = Array("Front Plane", "Top Plane", "Right Plane")
    Dim assyPlane As Object
    Dim compPlane As Object
    For Each swComp In comps
        ' A fixed component would be over-defined by the mates, so float it first
        If swComp.IsFixed Then
            swAssy.ClearSelection2 True
            boolstatus = swComp.Select4(False, Nothing, False)
            swAssy.UnfixComponent
        End If
        For i = 0 To 2
            Set assyPlane = findFeature(swAssy.FirstFeature, planeNames(i))
            Set compPlane = findFeature(swComp.FirstFeature, planeNames(i))
            If assyPlane Is Nothing Or compPlane Is Nothing Then
                MsgBox "Skipping " & planeNames(i) & " on " & swComp.Name2 & ", plane not found in the component or assembly."
            Else
                swAssy.ClearSelection2 True
                boolstatus = compPlane.Select2(False, 1)
                boolstatus = assyPlane.Select2(True, 1)
                Set swMate = swAssy.AddMate5(swMateType_e.swMateCOINCIDENT, swMateAlign_e.swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, longstatus)
                If swMate Is Nothing Then
                    MsgBox "Failed to mate " & planeNames(i) & " on " & swComp.Name2 & ", error " & longstatus
                End If
            End If
        Next i
    Next swComp

    swAssy.ClearSelection2 True
    swAssy.EditRebuild3

End Sub

' Returns the feature named featName found by walking the feature tree from swFeat, or Nothing
Function findFeature(ByVal swFeat As Object, featName As Variant) As Object
    Do While Not swFeat Is Nothing
        If swFeat.Name = featName Then
            Set findFeature = swFeat
            Exit Function
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
End Function
