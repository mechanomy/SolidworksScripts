Attribute VB_Name = "export3mfStepModule"
'Saves all bodies of the current part in STEP and 3MF formats
' Cobbled together from https://r1132100503382-eu1-3dswym.3dexperience.3ds.com/?ticket=ST-8891484-wWsd0ktdwh77ZH6Hokqm-cas#community:yUw32GbYTEqKdgY7-jbZPg/post:o-0ACu0uQo2VKG7uTDoWzQ
' Uses saveas https://help.solidworks.com/2022/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDocExtension~SaveAs3.html

'MIT License
'Copyright (c) 2023 Mechanomy
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


Dim swApp As Object

Dim swPart As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim path As String
Dim configName As String
Dim fname As String
Dim cnt As Integer
Dim Body_Vis_States() As Boolean
Dim bodyArray As Variant

Sub main()

    Set swApp = Application.SldWorks

    Set swPart = swApp.ActiveDoc

    Dim myModelView As Object
    Set myModelView = swPart.ActiveView
    myModelView.FrameState = swWindowState_e.swWindowMaximized

    ' Gets folder path of current part
    path = Left(swPart.GetPathName, InStrRev(swPart.GetPathName, "\") - 1)

    ' Get model configuration name
    configName = swPart.ConfigurationManager.ActiveConfiguration.Name

    ' Uncomment this line for STL files to use same file name as the part file
    fname = Left(swPart.GetTitle, InStrRev(swPart.GetTitle, ".") - 1) & "_" & configName

    ' creates an array of all the bodies in the current part
    bodyArray = swPart.GetBodies2(-1, False)

    ' Get current visibility state of all bodies, put into an array
    For cnt = 0 To UBound(bodyArray)
        Set swBody = bodyArray(cnt)
        If Not swBody Is Nothing Then
            ReDim Preserve Body_Vis_States(0 To cnt)
            Body_Vis_States(cnt) = swBody.Visible
            ' MsgBox ("Body " & cnt & " Visibility: " & Body_Vis_States(cnt))
        End If
    Next cnt

    ' Hide all bodies
    For cnt = 0 To UBound(bodyArray)
        Set swBody = bodyArray(cnt)
        If Not swBody Is Nothing Then
            swBody.HideBody (True)
            ' MsgBox ("Body " & cnt & " Hidden")
        End If
    Next cnt
    
    ' Show each body one by one, save as STL, then hide again
    For cnt = 0 To UBound(bodyArray)
        Set swBody = bodyArray(cnt)
        If Not swBody Is Nothing Then
            swBody.HideBody (False)

            longstatus = swPart.SaveAs3(path & "\" & fname & "_v" & cnt & ".3mf", 0, 2) 'save options enum https://help.solidworks.com/2022/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swSaveAsOptions_e.html
            ' MsgBox ("Body " & cnt & " Saved to 3mf")
            
            longstatus = swPart.SaveAs3(path & "\" & fname & "_v" & cnt & ".step", 0, 2) 'save as step for future
            ' MsgBox ("Body " & cnt & " Saved to STEP")

            swBody.HideBody (True)
        End If
    Next cnt

    ' Put bodies back in the original visibility state
    For cnt = 0 To UBound(bodyArray)
        Set swBody = bodyArray(cnt)
        If Not swBody Is Nothing Then
            swBody.HideBody (Not Body_Vis_States(cnt))
        End If
    Next cnt

    ' Open the window containing the STLs
    ' Shell "explorer.exe " & path & "\", vbNormalFocus

End Sub
