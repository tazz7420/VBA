Attribute VB_Name = "B_DrawLink"
Sub DrawLink()
Dim Xy3(0 To 3) As Double
Dim ObjCircle As AcadCircle
Dim LwObj As AcadLWPolyline
Dim TextObj As AcadText
Dim Xy2(0 To 2) As Double
Dim Ang As Double
Dim AcadDoc2 As AcadDocument
Dim AcadApp2 As New AcadApplication
Dim MM_Sheet(1 To 9) As Excel.Worksheet
Dim FixOrNot As Variant

Set AcadDoc2 = AcadApp.Documents.Open("C:\Users\WXZ\Desktop\Drawing1.dwg")
AcadDoc2.Application.Visible = True
AcadDoc2.WindowState = acMax

Set MM_Sheet(1) = ThisWorkbook.Worksheets("JUNCTIONS")
Set MM_Sheet(2) = ThisWorkbook.Worksheets("CONDUITS")
Set MM_Sheet(3) = ThisWorkbook.Worksheets("XSECTIONS")
Set MM_Sheet(4) = ThisWorkbook.Worksheets("TRANSECTS")
Set MM_Sheet(5) = ThisWorkbook.Worksheets("COORDINATES")
Set MM_Sheet(6) = ThisWorkbook.Worksheets("VERTICES")
Set MM_Sheet(7) = ThisWorkbook.Worksheets("Sheet7")
Set MM_Sheet(8) = ThisWorkbook.Worksheets("Sheet1")
Set MM_Sheet(9) = ThisWorkbook.Worksheets("Sheet2")



'k = 1
'Do While MM_Sheet(5).Cells(k, 1).Value <> ""
'    Xy2(0) = MM_Sheet(5).Cells(k, 2).Value
'    Xy2(1) = MM_Sheet(5).Cells(k, 3).Value
'    Xy2(2) = 0
'    Set TextObj = AcadDoc2.ModelSpace.AddText(MM_Sheet(5).Cells(k, 1).Value, Xy2, 1)
'    TextObj.Color = acBlue
'    TextObj.Update
'    Set ObjCircle = AcadDoc2.ModelSpace.AddCircle(Xy2, 0.5)
'    ObjCircle.Color = acBlue
'    ObjCircle.Update
'    k = k + 1
'    'AcadDoc2.SendCommand ("zoom" & vbCr & "c" & vbCr & Xy2(0) & "," & Xy2(1) & vbCr & "20" & vbCr)
'Loop


i = 1750
Do While MM_Sheet(2).Cells(i, 1).Value <> ""
    '道路中心線起始座標
    j = 1
    Do While MM_Sheet(3).Cells(j, 1).Value <> 0
        If MM_Sheet(2).Cells(i, 2).Value = MM_Sheet(5).Cells(j, 1).Value Then
            Xy3(0) = MM_Sheet(5).Cells(j, 2).Value
            Xy3(1) = MM_Sheet(5).Cells(j, 3).Value
            Exit Do
        End If
        j = j + 1
    Loop
    '---------------------
    
    '道路中心線終點座標
    j = 1
    Do While MM_Sheet(3).Cells(j, 1).Value <> 0
        If MM_Sheet(2).Cells(i, 3).Value = MM_Sheet(5).Cells(j, 1).Value Then
            Xy3(2) = MM_Sheet(5).Cells(j, 2).Value
            Xy3(3) = MM_Sheet(5).Cells(j, 3).Value
            Exit Do
        End If
        j = j + 1
    Loop
    '---------------------
    
    MM_Sheet(2).Activate
    MM_Sheet(2).Cells(i, 1).Select
    Ang = AzToAcadAngle(Pol(Xy3(1), Xy3(0), Xy3(3), Xy3(2)))
    Xy2(0) = Xy3(0) / 2 + Xy3(2) / 2
    Xy2(1) = Xy3(1) / 2 + Xy3(3) / 2
    Xy2(2) = 0
    Set LwObj = AcadDoc2.ModelSpace.AddLightWeightPolyline(Xy3)
    LwObj.Color = acRed
    LwObj.Update
    Set TextObj = AcadDoc2.ModelSpace.AddText(MM_Sheet(2).Cells(i, 1).Value, Xy2, 1)
    TextObj.Rotate Xy2, Ang
    TextObj.Color = acRed
    TextObj.Update
    AcadDoc2.SendCommand ("zoom" & vbCr & "c" & vbCr & Xy2(0) & "," & Xy2(1) & vbCr & "75" & vbCr)
    FixOrNot = InputBox("")
    If FixOrNot <> "" Then
        Stop
        i = i - 2
    End If
    On Error Resume Next
    LwObj.Delete
    TextObj.Delete
    i = i + 1
Loop
End Sub
