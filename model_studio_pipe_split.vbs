' Model Studio ActiveX VBScript template
' Splits selected pipe layout boundary into equal squares.
' Replace TODO API names according to your Model Studio COM model.

Option Explicit

Dim preferredSquareSide: preferredSquareSide = 1200
Dim splitLayerName: splitLayerName = "TRUCK_SPLIT_GRID"
Dim addLabels: addLabels = True

Sub Fail(msg)
  WScript.Echo "ERROR: " & msg
  WScript.Quit 1
End Sub

Function CeilDiv(a, b)
  CeilDiv = Int((a + b - 0.0000001) / b)
  If (a / b) > Int(a / b) Then CeilDiv = Int(a / b) + 1
End Function

Function ToLetter(idx)
  Dim i, m, s
  i = idx + 1
  s = ""
  Do While i > 0
    m = (i - 1) Mod 26
    s = Chr(65 + m) & s
    i = Int((i - 1) / 26)
  Loop
  ToLetter = s
End Function

Dim app
On Error Resume Next
Set app = CreateObject("ModelStudio.Application") ' TODO: exact ProgID
If Err.Number <> 0 Then
  Fail "Cannot create ActiveX object. " & Err.Description
End If
On Error GoTo 0

Dim doc
Set doc = app.ActiveDocument
If doc Is Nothing Then Fail "No active document."

Dim sel
Set sel = doc.Selection
If sel Is Nothing Then Fail "No selection found."
If sel.Count < 1 Then Fail "Select one closed contour first."

Dim boundary, bb
Set boundary = sel.Item(0)
Set bb = boundary.GetBoundingBox() ' expected: MinX, MinY, MaxX, MaxY

Dim minX, minY, maxX, maxY, width, height
minX = bb.MinX: minY = bb.MinY
maxX = bb.MaxX: maxY = bb.MaxY
width = maxX - minX
height = maxY - minY

If width <= 0 Or height <= 0 Then Fail "Boundary has invalid size."

Dim cols, rows, squareSide
cols = CeilDiv(width, preferredSquareSide)
rows = CeilDiv(height, preferredSquareSide)
squareSide = width / cols
If (height / rows) > squareSide Then squareSide = height / rows

cols = CeilDiv(width, squareSide)
rows = CeilDiv(height, squareSide)

Dim layer
Set layer = doc.Layers.GetOrCreate(splitLayerName) ' TODO

Dim c, r, x, y
For c = 0 To cols
  x = minX + c * squareSide
  doc.ModelSpace.AddLine x, minY, x, maxY, layer ' TODO
Next

For r = 0 To rows
  y = minY + r * squareSide
  doc.ModelSpace.AddLine minX, y, maxX, y, layer ' TODO
Next

If addLabels Then
  Dim cx, cy, label
  For r = 0 To rows - 1
    For c = 0 To cols - 1
      cx = minX + (c + 0.5) * squareSide
      cy = minY + (r + 0.5) * squareSide
      label = ToLetter(r) & CStr(c + 1)
      doc.ModelSpace.AddText label, cx, cy, squareSide * 0.12, layer ' TODO
    Next
  Next
End If

doc.Regen

WScript.Echo "Done. Grid created: " & rows & " x " & cols & _
             ", square side = " & Round(squareSide, 2)
