Attribute VB_Name = "C_FindDh"
Sub FindDh()
Dim Xy3(0 To 3) As Double
Dim LwObj As AcadLWPolyline
Dim TextObj As AcadText
Dim Xy2(0 To 2) As Double
Dim Ang As Double
Dim aintFilterType(0) As Integer
Dim avarFilterValue(0) As Variant
Dim objAcadSelectionSet As AcadSelectionSet
Dim SelectPoint1 As Variant
Dim SelectPoint2 As Variant
Dim objAcadEntity As AcadEntity
Dim LwObj1 As AcadLWPolyline
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



i = 1

Do While M_Sheet(2).Cells(i, 1).Value <> ""
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
    AcadDoc.SendCommand ("zoom" & vbCr & "c" & vbCr & ZoomPoint1(0) & "," & ZoomPoint1(1) & vbCr & "200" & vbCr)
    On Error Resume Next
    AcadDoc.SelectionSets.Item("TestSelectionSetFilter").Delete
    Set objAcadSelectionSet = AcadDoc.SelectionSets.Add("TestSelectionSetFilter")
    'objAcadSelectionSet.Select acSelectionSetCrossing, SelectPoint1, SelectPoint2
    Set LwObj00 = AcadDoc.ModelSpace.AddLightWeightPolyline(Xy3)
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
    LwObj03.Update
    Call RbtSelectCrossing(objAcadSelectionSet, SelectPoint1, SelectPoint2, 2, 0, "LWPOLYLINE", 8, "道路斷面")
    'MsgBox (objAcadSelectionSet.Count)
    

    '抓斷面--------------------------------------
    n = 0
    For Each objAcadEntity In objAcadSelectionSet
        If objAcadEntity.Layer = "道路斷面" Then
            Set LwObj1 = objAcadEntity
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
    
    If SearchDist = 0 Then
        SearchDist = 10
    End If
    
    '道路中心線起始座標
    j = 1
    Do While M_Sheet(3).Cells(j, 1).Value <> 0
        If M_Sheet(2).Cells(i, 2).Value = M_Sheet(5).Cells(j, 1).Value Then
            Xy3(0) = M_Sheet(5).Cells(j, 2).Value
            Xy3(1) = M_Sheet(5).Cells(j, 3).Value
            Exit Do
        End If
        j = j + 1
    Loop
    '---------------------
    
    '道路中心線終點座標
    j = 1
    Do While M_Sheet(3).Cells(j, 1).Value <> 0
        If M_Sheet(2).Cells(i, 3).Value = M_Sheet(5).Cells(j, 1).Value Then
            Xy3(2) = M_Sheet(5).Cells(j, 2).Value
            Xy3(3) = M_Sheet(5).Cells(j, 3).Value
            Exit Do
        End If
        j = j + 1
    Loop
    '---------------------
    Ang = AzToAcadAngle(Pol(Xy3(1), Xy3(0), Xy3(3), Xy3(2)))
    '選取側溝-----------------------------
    AcadDoc.SelectionSets.Item("DhSelectionSetFilter").Delete
    Set objAcadSelectionSet = AcadDoc.SelectionSets.Add("DhSelectionSetFilter")
    'objAcadSelectionSet.Select acSelectionSetCrossing, SelectPoint1, SelectPoint2
    Xy2(0) = Xy3(0)
    Xy2(1) = Xy3(1)
    Xy2(2) = 0
    SelectPoint1 = Xy2
    SelectPoint2 = AcadDoc.Utility.PolarPoint(SelectPoint1, Ang + Pi / 2, SearchDist)
    Set LineObj = AcadDoc.ModelSpace.AddLine(SelectPoint1, SelectPoint2)
    LineObj.Color = acBlue
    LineObj.Update
    Call RbtSelectCrossing(objAcadSelectionSet, SelectPoint1, SelectPoint2, 2, 0, "LWPOLYLINE", 8, "側溝測線")
    ZoomPoint1(0) = SelectPoint2(0)
    ZoomPoint1(1) = SelectPoint2(1)
    ZoomPoint1(2) = 0
    AcadDoc.SendCommand ("zoom" & vbCr & "c" & vbCr & ZoomPoint1(0) & "," & ZoomPoint1(1) & vbCr & "200" & vbCr)
'    newHour = Hour(Now())
'    newMinute = Minute(Now())
'    newSecond = Second(Now()) + 1
'    waitTime = TimeSerial(newHour, newMinute, newSecond)
'    Application.Wait waitTime
    Xy2(0) = Xy3(0)
    Xy2(1) = Xy3(1)
    Xy2(2) = 0
    SelectPoint1 = Xy2
    SelectPoint2 = AcadDoc.Utility.PolarPoint(SelectPoint1, Ang - Pi / 2, SearchDist)
    Set LineObj1 = AcadDoc.ModelSpace.AddLine(SelectPoint1, SelectPoint2)
    LineObj1.Color = acBlue
    LineObj1.Update
    Call RbtSelectCrossing(objAcadSelectionSet, SelectPoint1, SelectPoint2, 2, 0, "LWPOLYLINE", 8, "側溝測線")
    Xy2(0) = Xy3(2) / 2 + Xy3(0) / 2
    Xy2(1) = Xy3(3) / 2 + Xy3(1) / 2
    Xy2(2) = 0
    SelectPoint1 = Xy2
    SelectPoint2 = AcadDoc.Utility.PolarPoint(SelectPoint1, Ang + Pi / 2, SearchDist)
    Set LineObj2 = AcadDoc.ModelSpace.AddLine(SelectPoint1, SelectPoint2)
    LineObj2.Color = acBlue
    LineObj2.Update
    Call RbtSelectCrossing(objAcadSelectionSet, SelectPoint1, SelectPoint2, 2, 0, "LWPOLYLINE", 8, "側溝測線")
    Xy2(0) = Xy3(2) / 2 + Xy3(0) / 2
    Xy2(1) = Xy3(3) / 2 + Xy3(1) / 2
    Xy2(2) = 0
    SelectPoint1 = Xy2
    SelectPoint2 = AcadDoc.Utility.PolarPoint(SelectPoint1, Ang - Pi / 2, SearchDist)
    Set LineObj3 = AcadDoc.ModelSpace.AddLine(SelectPoint1, SelectPoint2)
    LineObj3.Color = acBlue
    LineObj3.Update
    Call RbtSelectCrossing(objAcadSelectionSet, SelectPoint1, SelectPoint2, 2, 0, "LWPOLYLINE", 8, "側溝測線")
    Xy2(0) = Xy3(2)
    Xy2(1) = Xy3(3)
    Xy2(2) = 0
    SelectPoint1 = Xy2
    SelectPoint2 = AcadDoc.Utility.PolarPoint(SelectPoint1, Ang + Pi / 2, SearchDist)
    Set LineObj4 = AcadDoc.ModelSpace.AddLine(SelectPoint1, SelectPoint2)
    LineObj4.Color = acBlue
    LineObj4.Update
    Call RbtSelectCrossing(objAcadSelectionSet, SelectPoint1, SelectPoint2, 2, 0, "LWPOLYLINE", 8, "側溝測線")
    Xy2(0) = Xy3(2)
    Xy2(1) = Xy3(3)
    Xy2(2) = 0
    SelectPoint1 = Xy2
    SelectPoint2 = AcadDoc.Utility.PolarPoint(SelectPoint1, Ang - Pi / 2, SearchDist)
    Set LineObj5 = AcadDoc.ModelSpace.AddLine(SelectPoint1, SelectPoint2)
    LineObj5.Color = acBlue
    LineObj5.Update
    Call RbtSelectCrossing(objAcadSelectionSet, SelectPoint1, SelectPoint2, 2, 0, "LWPOLYLINE", 8, "側溝測線")
    'MsgBox (objAcadSelectionSet.Count)
    ZoomPoint2(0) = SelectPoint2(0)
    ZoomPoint2(1) = SelectPoint2(1)
    ZoomPoint2(2) = 0
    AcadDoc.SendCommand ("zoom" & vbCr & "c" & vbCr & ZoomPoint1(0) & "," & ZoomPoint1(1) & vbCr & "200" & vbCr)
    newHour = Hour(Now())
    newMinute = Minute(Now())
    newSecond = Second(Now()) + 1
    waitTime = TimeSerial(newHour, newMinute, newSecond)
    Application.Wait waitTime
    AcadDoc.SendCommand ("zoom" & vbCr & "c" & vbCr & ZoomPoint1(0) & "," & ZoomPoint1(1) & vbCr & "200" & vbCr)
    newHour = Hour(Now())
    newMinute = Minute(Now())
    newSecond = Second(Now()) + 1
    waitTime = TimeSerial(newHour, newMinute, newSecond)
    Application.Wait waitTime

    qq = 13
    For Each objAcadEntity In objAcadSelectionSet
        Set LwObj1 = objAcadEntity
        LineReturn = LineObj.IntersectWith(LwObj1, acExtendNone)
        LineReturn1 = LineObj1.IntersectWith(LwObj1, acExtendNone)
        LineReturn2 = LineObj2.IntersectWith(LwObj1, acExtendNone)
        LineReturn3 = LineObj3.IntersectWith(LwObj1, acExtendNone)
        LineReturn4 = LineObj4.IntersectWith(LwObj1, acExtendNone)
        LineReturn5 = LineObj5.IntersectWith(LwObj1, acExtendNone)
        If UBound(LineReturn) >= 2 Or UBound(LineReturn1) >= 2 Or UBound(LineReturn2) >= 2 Or UBound(LineReturn3) >= 2 Or UBound(LineReturn4) >= 2 Or UBound(LineReturn5) >= 2 Then
            LwObj1.Color = acRed
            LwObj1.Update
            StartP(0) = LwObj1.Coordinates(0)
            StartP(1) = LwObj1.Coordinates(1)
            EndP(0) = LwObj1.Coordinates(UBound(LwObj1.Coordinates) - 1)
            EndP(1) = LwObj1.Coordinates(UBound(LwObj1.Coordinates))
            q = 2
            
            Do While M_Sheet(7).Cells(q, 1).Value <> 0
                If Abs(StartP(0) - M_Sheet(7).Cells(q, 3).Value) < 3 And Abs(StartP(1) - M_Sheet(7).Cells(q, 4).Value) < 3 And Abs(EndP(0) - M_Sheet(7).Cells(q, 6).Value) < 3 And Abs(EndP(1) - M_Sheet(7).Cells(q, 7).Value) < 3 Then
                    M_Sheet(2).Cells(i, qq).Value = M_Sheet(7).Cells(q, 2).Value
                    qq = qq + 1
                ElseIf Abs(StartP(0) - M_Sheet(7).Cells(q, 6).Value) < 3 And Abs(StartP(1) - M_Sheet(7).Cells(q, 7).Value) < 3 And Abs(EndP(0) - M_Sheet(7).Cells(q, 3).Value) < 3 And Abs(EndP(1) - M_Sheet(7).Cells(q, 4).Value) < 3 Then
                    M_Sheet(2).Cells(i, qq).Value = M_Sheet(7).Cells(q, 2).Value
                    qq = qq + 1
                End If
                q = q + 1
            Loop
        End If
    Next
    
    
    'DhFindOK = ""
    'DhFindOK = CStr(InputBox("側溝OK嗎"))

    'M_Sheet(2).Cells(i, 10).Value = DhFindOK
    
    For Each objAcadEntity In objAcadSelectionSet
        Set LwObj1 = objAcadEntity
        LwObj1.Color = acByLayer
        LwObj1.Update
    Next
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

    'MsgBox (objAcadSelectionSet.Count)
    
    i = i + 1
Loop

End Sub
