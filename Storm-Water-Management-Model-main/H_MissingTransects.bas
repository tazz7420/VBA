Attribute VB_Name = "H_MissingTransects"
'Public AcadApp As New AcadApplication
'Public AcadDoc As AcadDocument
'Public M_Sheet(1 To 9) As Excel.Worksheet

Sub MissingTransects()
Dim Xy3(0 To 3) As Double
Dim LwObj As AcadLWPolyline
Dim TextObj As AcadText
Dim Xy2(0 To 2) As Double
Dim ObjRoadLine As AcadLWPolyline
Dim Ang As Double
Dim aintFilterType(0) As Integer
Dim avarFilterValue(0) As Variant
Dim objAcadSelectionSet As AcadSelectionSet
Dim SelectPoint1 As Variant
Dim SelectPoint2 As Variant
Dim objAcadEntity As AcadEntity
Dim LwObj1 As AcadLWPolyline
Dim LwObj11 As AcadLWPolyline
Dim LwObj00 As AcadLWPolyline
Dim LwObj01 As AcadLWPolyline
Dim LwObj02 As AcadLWPolyline
Dim LwObj03 As AcadLWPolyline
Dim LwObj2 As AcadLWPolyline
Dim varReturn As Variant
Dim varReturn1 As Variant
Dim varReturn2 As Variant
Dim varReturn3 As Variant
Dim varReturn4 As Variant
Dim varReturn5 As Variant
Dim SearchDist As Double
Dim LineObj As AcadLine
Dim LineObj1 As AcadLine
Dim LineObj2 As AcadLine
Dim LineObj3 As AcadLine
Dim LineObj4 As AcadLine
Dim LineObj5 As AcadLine
Dim LineReturn As Variant
Dim LineReturn1 As Variant
Dim LineReturn2 As Variant
Dim LineReturn3 As Variant
Dim LineReturn4 As Variant
Dim LineReturn5 As Variant
Dim StartP(0 To 1), EndP(0 To 1) As Double
Dim ZoomPoint1(0 To 2) As Double
Dim ZoomPoint2(0 To 2) As Double
Dim SelectPoint3 As Variant
Dim Transects() As Variant
Dim TransectsDist As Double
Dim AvgEL As Double
Dim AvgWidth As Double
Dim AvgDepth As Double
Dim FixDist, FixDist1 As Double
Dim DhCount, Dh, Dh0, Dh00, Dh1, Dh2 As Integer
Dim HaveDh As Integer
Dim Depth1, Depth2 As Double
Dim ObjCircle As AcadCircle

Set AcadDoc = AcadApp.Documents.Open("C:\Users\WXZ\Desktop\Drawing6.dwg")
AcadDoc.Application.Visible = True
AcadDoc.WindowState = acMax

Set M_Sheet(1) = ThisWorkbook.Worksheets("JUNCTIONS")
Set M_Sheet(2) = ThisWorkbook.Worksheets("CONDUITS")
Set M_Sheet(3) = ThisWorkbook.Worksheets("XSECTIONS")
Set M_Sheet(4) = ThisWorkbook.Worksheets("TRANSECTS")
Set M_Sheet(5) = ThisWorkbook.Worksheets("COORDINATES")
Set M_Sheet(6) = ThisWorkbook.Worksheets("VERTICES")
Set M_Sheet(7) = ThisWorkbook.Worksheets("Sheet7")
Set M_Sheet(8) = ThisWorkbook.Worksheets("Sheet1")
Set M_Sheet(9) = ThisWorkbook.Worksheets("Sheet2")

i = 3347
w = 1
j = 9613
Do While M_Sheet(2).Cells(i, 1).Value <> ""
    k = 1
    M_Sheet(2).Activate
    M_Sheet(2).Cells(i, 1).Select
    u = 2
    Xy3(0) = CDbl(M_Sheet(6).Cells(i, 2).Value)
    Xy3(1) = CDbl(M_Sheet(6).Cells(i, 3).Value)
    Xy3(2) = CDbl(M_Sheet(6).Cells(i, 2).Value - 1)
    Xy3(3) = CDbl(M_Sheet(6).Cells(i, 3).Value + 1)
    Xy2(0) = Xy3(0)
    Xy2(1) = Xy3(1)
    Xy2(2) = 0
    SelectPoint1 = Xy2
    Xy2(0) = Xy3(2)
    Xy2(1) = Xy3(3)
    Xy2(2) = 0
    SelectPoint2 = Xy2
    ZoomPoint1(0) = Xy3(0)
    ZoomPoint1(1) = Xy3(1)
    ZoomPoint1(2) = 0
    AcadDoc.SendCommand ("zoom" & vbCr & "c" & vbCr & ZoomPoint1(0) & "," & ZoomPoint1(1) & vbCr & "50" & vbCr)
    On Error Resume Next
    AcadDoc.SelectionSets.Item("TestSelectionSetFilter").Delete
    Set objAcadSelectionSet = AcadDoc.SelectionSets.Add("TestSelectionSetFilter")
    'objAcadSelectionSet.Select acSelectionSetCrossing, SelectPoint1, SelectPoint2
    Set LwObj00 = AcadDoc.ModelSpace.AddLightWeightPolyline(Xy3)
    LwObj00.Layer = "0"
    LwObj00.Update
    Call RbtSelectCrossing(objAcadSelectionSet, SelectPoint1, SelectPoint2, 2, 0, "LWPOLYLINE", 8, "道路斷面")
    'MsgBox (objAcadSelectionSet.Count)
    
    
    Xy3(0) = CDbl(M_Sheet(6).Cells(i, 2).Value)
    Xy3(1) = CDbl(M_Sheet(6).Cells(i, 3).Value)
    Xy3(2) = CDbl(M_Sheet(6).Cells(i, 2).Value + 1)
    Xy3(3) = CDbl(M_Sheet(6).Cells(i, 3).Value + 1)
    Xy2(0) = Xy3(0)
    Xy2(1) = Xy3(1)
    Xy2(2) = 0
    SelectPoint1 = Xy2
    Xy2(0) = Xy3(2)
    Xy2(1) = Xy3(3)
    Xy2(2) = 0
    SelectPoint2 = Xy2
    Set LwObj01 = AcadDoc.ModelSpace.AddLightWeightPolyline(Xy3)
    LwObj01.Layer = "0"
    LwObj01.Update
    Call RbtSelectCrossing(objAcadSelectionSet, SelectPoint1, SelectPoint2, 2, 0, "LWPOLYLINE", 8, "道路斷面")
    
    Xy3(0) = CDbl(M_Sheet(6).Cells(i, 2).Value)
    Xy3(1) = CDbl(M_Sheet(6).Cells(i, 3).Value)
    Xy3(2) = CDbl(M_Sheet(6).Cells(i, 2).Value + 1)
    Xy3(3) = CDbl(M_Sheet(6).Cells(i, 3).Value - 1)
    Xy2(0) = Xy3(0)
    Xy2(1) = Xy3(1)
    Xy2(2) = 0
    SelectPoint1 = Xy2
    Xy2(0) = Xy3(2)
    Xy2(1) = Xy3(3)
    Xy2(2) = 0
    SelectPoint2 = Xy2
    Set LwObj02 = AcadDoc.ModelSpace.AddLightWeightPolyline(Xy3)
    LwObj02.Layer = "0"
    LwObj02.Update
    Call RbtSelectCrossing(objAcadSelectionSet, SelectPoint1, SelectPoint2, 2, 0, "LWPOLYLINE", 8, "道路斷面")
    
    Xy3(0) = CDbl(M_Sheet(6).Cells(i, 2).Value)
    Xy3(1) = CDbl(M_Sheet(6).Cells(i, 3).Value)
    Xy3(2) = CDbl(M_Sheet(6).Cells(i, 2).Value - 1)
    Xy3(3) = CDbl(M_Sheet(6).Cells(i, 3).Value - 1)
    Xy2(0) = Xy3(0)
    Xy2(1) = Xy3(1)
    Xy2(2) = 0
    SelectPoint1 = Xy2
    Xy2(0) = Xy3(2)
    Xy2(1) = Xy3(3)
    Xy2(2) = 0
    SelectPoint2 = Xy2
    Set LwObj03 = AcadDoc.ModelSpace.AddLightWeightPolyline(Xy3)
    LwObj03.Layer = "0"
    LwObj03.Update
    Call RbtSelectCrossing(objAcadSelectionSet, SelectPoint1, SelectPoint2, 2, 0, "LWPOLYLINE", 8, "道路斷面")
    'MsgBox (objAcadSelectionSet.Count)
    

    '抓斷面--------------------------------------
    n = 0
    For Each objAcadEntity In objAcadSelectionSet
        If objAcadEntity.Layer = "道路斷面" Then
            Set LwObj1 = objAcadEntity
            Set LwObj11 = objAcadEntity
            varReturn = LwObj00.IntersectWith(LwObj1, acExtendNone)
            varReturn1 = LwObj01.IntersectWith(LwObj1, acExtendNone)
            varReturn2 = LwObj02.IntersectWith(LwObj1, acExtendNone)
            varReturn3 = LwObj03.IntersectWith(LwObj1, acExtendNone)
            If UBound(varReturn) >= 2 Or UBound(varReturn1) >= 2 Or UBound(varReturn2) >= 2 Or UBound(varReturn3) >= 2 Then
                n = UBound(LwObj1.Coordinates)
            End If
        End If
    Next
    SearchDist = Hdist(LwObj1.Coordinates(0), LwObj1.Coordinates(1), LwObj1.Coordinates(UBound(LwObj1.Coordinates) - 1), LwObj1.Coordinates(UBound(LwObj1.Coordinates))) / 2 + 10
    '------------------------------------------------------
    

    'MsgBox (objAcadSelectionSet.Count)
    If n = 0 Then
        XXXXX = MsgBox("")
        M_Sheet(8).Cells(j, 1).Value = M_Sheet(8).Cells(j - 1, 1).Value + 1
        M_Sheet(8).Cells(j, 2).Value = CDbl(M_Sheet(6).Cells(i, 2).Value)
        M_Sheet(8).Cells(j, 3).Value = CDbl(M_Sheet(6).Cells(i, 3).Value)
        Do While M_Sheet(1).Cells(k, 1).Value <> ""
            If M_Sheet(1).Cells(k, 1).Value = M_Sheet(2).Cells(i, 2).Value Then
                Depth1 = M_Sheet(1).Cells(k, 2).Value + M_Sheet(1).Cells(k, 3).Value
            End If
            If M_Sheet(1).Cells(k, 1).Value = M_Sheet(2).Cells(i, 3).Value Then
                Depth2 = M_Sheet(1).Cells(k, 2).Value + M_Sheet(1).Cells(k, 3).Value
            End If
            k = k + 1
        Loop
        M_Sheet(8).Activate
        M_Sheet(8).Cells(j, 4).Select
        M_Sheet(8).Cells(j, 4).Value = Depth1 / 2 + Depth2 / 2
        M_Sheet(8).Cells(j, 7).Value = Depth1 / 2 + Depth2 / 2 - 0.1
        M_Sheet(8).Cells(j, 10).Value = Depth1 / 2 + Depth2 / 2 - 0.1
        objAcadSelectionSet.SelectOnScreen
        jj = 5
        For Each objAcadEntity In objAcadSelectionSet
            If objAcadEntity.Layer = "0" Then
                Set ObjCircle = objAcadEntity
                M_Sheet(8).Cells(j, jj).Value = ObjCircle.Center(0)
                M_Sheet(8).Cells(j, jj + 1).Value = ObjCircle.Center(1)
                jj = jj + 3
                ObjCircle.Delete
            End If
        Next
        j = j + 1
        
        GoTo Z
    End If
    
Z:
    For Each objAcadEntity In objAcadSelectionSet
        Set LwObj1 = objAcadEntity
        LwObj1.Color = acByLayer
        LwObj1.Update
    Next
    'M_Sheet(2).Cells(i, 11) = DhCount
    DhCount = 0
    i = i + 1
    LwObj00.Delete
    LwObj01.Delete
    LwObj02.Delete
    LwObj03.Delete
    LineObj.Delete
    LineObj1.Delete
    LineObj2.Delete
    LineObj3.Delete
    LineObj4.Delete
    LineObj5.Delete
    LineObj5.Update
Loop

End Sub
