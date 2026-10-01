Attribute VB_Name = "saveIsometricPng"
'One-button save png to file.
' PNG has the same name as the part file.
' Taken from the isometric perspective.

'MIT License
'Copyright (c) 2023 Mechanomy
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


Dim swApp As Object

Dim swPart As Object
Dim partPath As String
Dim pngPath As String
Dim viewportPath As String
Dim largePath As String
Dim swView As Object
Dim viewOrientation As Object
Dim viewTranslation As Object
Dim viewScale As Double
Dim viewportWidth As Long
Dim viewportHeight As Long

' SaveBMP ignores the display anti-aliasing setting, so render this many times the viewport size and downscale to smooth the edges
Const supersample As Long = 3

Sub main()

  Set swApp = Application.SldWorks

  Set swPart = swApp.ActiveDoc
  If swPart Is Nothing Then Exit Sub

  ' The PNG goes next to the part file, so the part must be saved first
  partPath = swPart.GetPathName
  If partPath = "" Then
    MsgBox "Save the part before running saveIsometricPng."
    Exit Sub
  End If
  pngPath = Left(partPath, InStrRev(partPath, ".")) & "png"
  viewportPath = Left(partPath, InStrRev(partPath, ".")) & "viewport.bmp"
  largePath = Left(partPath, InStrRev(partPath, ".")) & "supersample.bmp"

  ' Remember the current view so it can be restored after saving
  Set swView = swPart.ActiveView
  Set viewOrientation = swView.Orientation3
  Set viewTranslation = swView.Translation3
  viewScale = swView.Scale2

  swPart.ShowNamedView2 "*Isometric", swStandardViews_e.swIsometricView
  swPart.ViewZoomtofit2

  ' SaveAs writes PNGs with alpha = 0 on every pixel, so they display as blank; render BMPs with SaveBMP and convert those instead.
  ' FrameWidth/FrameHeight include window borders, so a SaveBMP at size 0 (the viewport) gives the exact graphics area size.
  If Not swPart.SaveBMP(viewportPath, 0, 0) Then
    MsgBox "Failed to save " & viewportPath
  Else
    readBmpSize viewportPath, viewportWidth, viewportHeight
    Kill viewportPath
    If Not swPart.SaveBMP(largePath, viewportWidth * supersample, viewportHeight * supersample) Then
      MsgBox "Failed to save " & largePath
    ElseIf Not downscaleToPng(largePath, pngPath, viewportWidth, viewportHeight) Then
      MsgBox "Failed to convert " & largePath & " to " & pngPath ' the large render is kept for inspection
    Else
      Kill largePath
    End If
  End If

  swView.Orientation3 = viewOrientation
  swView.Translation3 = viewTranslation
  swView.Scale2 = viewScale
  swView.GraphicsRedraw Nothing

End Sub

' Reads the pixel size from a BMP header: width and height are 32-bit little-endian at byte offsets 18 and 22
Sub readBmpSize(path As String, width As Long, height As Long)
  Dim fileNum As Integer
  fileNum = FreeFile
  Open path For Binary Access Read As #fileNum
  Get #fileNum, 19, width ' Get positions are 1-based
  Get #fileNum, 23, height
  Close #fileNum
  height = Abs(height) ' negative height marks a top-down BMP
End Sub

' Resizes srcPath to width x height with high-quality bicubic filtering and saves it as a PNG, using .NET System.Drawing via PowerShell
Function downscaleToPng(srcPath As String, dstPath As String, width As Long, height As Long) As Boolean
  Dim script As String
  Dim exitCode As Long

  ' Single quotes delimit the paths in PowerShell, so double any in the paths
  script = "Add-Type -AssemblyName System.Drawing; " & _
    "$src = [Drawing.Image]::FromFile('" & Replace(srcPath, "'", "''") & "'); " & _
    "$dst = New-Object Drawing.Bitmap " & width & ", " & height & ", ([Drawing.Imaging.PixelFormat]::Format24bppRgb); " & _
    "$g = [Drawing.Graphics]::FromImage($dst); " & _
    "$g.InterpolationMode = 'HighQualityBicubic'; " & _
    "$g.PixelOffsetMode = 'HighQuality'; " & _
    "$g.SmoothingMode = 'HighQuality'; " & _
    "$g.DrawImage($src, 0, 0, " & width & ", " & height & "); " & _
    "$src.Dispose(); " & _
    "$dst.Save('" & Replace(dstPath, "'", "''") & "', [Drawing.Imaging.ImageFormat]::Png); " & _
    "$g.Dispose(); $dst.Dispose()"

  ' Window style 0 hides the console; True waits for PowerShell to finish
  exitCode = CreateObject("WScript.Shell").Run("powershell -NoProfile -ExecutionPolicy Bypass -Command """ & script & """", 0, True)
  downscaleToPng = (exitCode = 0)
End Function
