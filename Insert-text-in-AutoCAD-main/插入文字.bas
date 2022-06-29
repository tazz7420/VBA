Attribute VB_Name = "插入文字"
Public Sub 插入文字()
Dim AcadApp As New AcadApplication
Dim AcadDoc As AcadDocument
Dim Datprt As Excel.Range
Dim DatCount As Integer
Dim i As Integer
Dim FileName As String
Dim CoorX, CoorY, TextSize, PointSize, MoveX, MoveY As Double
Dim TextName, LayerName As String
Dim TextColor, PointColor As Integer
Dim TextPt(0 To 2) As Double
Dim PointPt(0 To 2) As Double
Dim AcadText As AcadText
Dim AcadPoint As AcadPoint
Dim PointType As Integer



'計算檢查點數量
Set Datprt = ThisWorkbook.Worksheets("Sheet1").Cells(1, 1).CurrentRegion
DatCount = Datprt.EntireRow.Count - 1
'MsgBox (DatCount)
If DatCount = 1 Then
    MsgBox ("Null")
    GoTo Z
End If

FileName = ThisWorkbook.Path & "/插入文字.dwg"
If ThisWorkbook.Worksheets("Sheet1").Cells(i + 1, 1) = 0 Then
    GoTo Z
End If
On Error Resume Next
Set AcadDoc = AcadApp.Documents.Open(FileName)
AcadDoc.Application.Visible = True
AcadDoc.WindowState = acMax

PointSize = Val(ThisWorkbook.Worksheets("Sheet1").Cells(2, 8).Value)
PointType = CInt(ThisWorkbook.Worksheets("Sheet1").Cells(2, 7).Value)
AcadDoc.SetVariable "PDMODE", PointType
AcadDoc.SetVariable "PDSIZE", PointSize



For i = 1 To DatCount
  
    
    'Start Drawing
    CoorX = Val(ThisWorkbook.Worksheets("Sheet1").Cells(i + 1, 3).Value)
    CoorY = Val(ThisWorkbook.Worksheets("Sheet1").Cells(i + 1, 2).Value)
    TextName = ThisWorkbook.Worksheets("Sheet1").Cells(i + 1, 4).Value
    TextSize = Val(ThisWorkbook.Worksheets("Sheet1").Cells(i + 1, 5).Value)
    TextColor = CInt(ThisWorkbook.Worksheets("Sheet1").Cells(i + 1, 6).Value)
    PointColor = CInt(ThisWorkbook.Worksheets("Sheet1").Cells(i + 1, 9).Value)
    LayerName = ThisWorkbook.Worksheets("Sheet1").Cells(i + 1, 10).Value
    MoveX = Val(ThisWorkbook.Worksheets("Sheet1").Cells(i + 1, 11).Value)
    MoveY = Val(ThisWorkbook.Worksheets("Sheet1").Cells(i + 1, 12).Value)
    AcadDoc.Layers.Add (LayerName)
    TextPt(0) = CoorX + MoveX
    TextPt(1) = CoorY + MoveY
    TextPt(2) = 0
    Set AcadText = AcadDoc.ModelSpace.AddText(TextName, TextPt, TextSize)
    AcadText.Alignment = acAlignmentMiddleCenter
    AcadText.TextAlignmentPoint = TextPt
    AcadText.Color = TextColor
    AcadText.Layer = LayerName
    AcadText.Update
    PointPt(0) = CoorX
    PointPt(1) = CoorY
    PointPt(2) = 0
    Set AcadPoint = AcadDoc.ModelSpace.AddPoint(PointPt)
    AcadPoint.Color = PointColor
    AcadPoint.Layer = LayerName
Next
    
Z:
AcadDoc.SaveAs ThisWorkbook.Path & "\Text", acR12_dxf
AcadDoc.Close False
MsgBox ("Done")
End Sub
