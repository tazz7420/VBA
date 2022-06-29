Attribute VB_Name = "C_FindDhTf"
Sub FindDhTf()
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
Dim Interpolation, Interpolation1, Interpolation2 As Double
Dim ObjCircle(0 To 99) As AcadCircle
Dim ObjLine(0 To 99) As AcadLine
Dim CircleNum, PointNum, LineNum As Integer
Dim Xy22(0 To 2) As Double
Dim Xy222(0 To 2) As Double
Dim FixOrNot, FixOrNot1 As Integer
Dim FixI, FixJ As Integer
Dim FixEL1, FixEL2 As Double




i = InputBox("欲開始CONDUITS的行數")
w = InputBox("欲開始TRANSECTS的行數")
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
    AcadDoc.SendCommand ("zoom" & vbCr & "c" & vbCr & ZoomPoint1(0) & "," & ZoomPoint1(1) & vbCr & "50" & vbCr)
    On Error Resume Next
    AcadDoc.SelectionSets.Item("TestSelectionSetFilter").Delete
    Set objAcadSelectionSet = AcadDoc.SelectionSets.Add("TestSelectionSetFilter")
    'objAcadSelectionSet.Select acSelectionSetCrossing, SelectPoint1, SelectPoint2
    Set LwObj00 = AcadDoc.ModelSpace.AddLightWeightPolyline(Xy3)
    LwObj00.Lineweight = acLnWt100
    LwObj00.Update
    LwObj00.Color = acRed
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
    LwObj01.Lineweight = acLnWt100
    LwObj01.Update
    LwObj01.Color = acRed
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
    LwObj02.Lineweight = acLnWt100
    LwObj02.Update
    LwObj02.Color = acRed
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
    LwObj03.Lineweight = acLnWt100
    LwObj03.Color = acRed
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
    Set ObjRoadLine = AcadDoc.ModelSpace.AddLightWeightPolyline(Xy3)
    ObjRoadLine.Color = acRed
    ObjRoadLine.Update
    
    Ang = AzToAcadAngle(Pol(Xy3(1), Xy3(0), Xy3(3), Xy3(2)))
    '選取側溝-----------------------------
    AcadDoc.SelectionSets.Item("DhSelectionSetFilter").Delete
    Set objAcadSelectionSet = AcadDoc.SelectionSets.Add("DhSelectionSetFilter")
    AcadDoc.Activate
    'objAcadSelectionSet.Select acSelectionSetCrossing, SelectPoint1, SelectPoint2
    objAcadSelectionSet.SelectOnScreen
    For Each objAcadEntity In objAcadSelectionSet
        If objAcadEntity.Layer = "側溝測線" Then
            Set LwObj2 = objAcadEntity
            LwObj2.Color = acRed
        End If
    Next
'    newHour = Hour(Now())
'    newMinute = Minute(Now())
'    newSecond = Second(Now()) + 1
'    waitTime = TimeSerial(newHour, newMinute, newSecond)
'    Application.Wait waitTime



    
    
    'DhFindOK = ""
    'DhFindOK = CStr(InputBox("側溝OK嗎"))

    'M_Sheet(2).Cells(i, 10).Value = DhFindOK
    
'    For Each objAcadEntity In objAcadSelectionSet
'        Set LwObj1 = objAcadEntity
'        LwObj1.Color = acByLayer
'        LwObj1.Update
'    Next

    'MsgBox (objAcadSelectionSet.Count)
    If n = 0 Then
        XXXXX = MsgBox("")
        Stop
    End If
    q = 2
    Xy3(0) = CDbl(M_Sheet(6).Cells(i, 2).Value)
    Xy3(1) = CDbl(M_Sheet(6).Cells(i, 3).Value)
    '抓側溝--------------------------------------
    m = 0
    Ang = AzToAcadAngle(Pol(LwObj11.Coordinates(1), LwObj11.Coordinates(0), Xy3(1), Xy3(0)))
    '抓起始側側溝--------------------------------------
    Set objAcadSelectionSet = AcadDoc.SelectionSets.Add("DhSelectionSetFilter")
    'MsgBox (objAcadSelectionSet.Count)
    Xy2(0) = Xy3(0)
    Xy2(1) = Xy3(1)
    Xy2(2) = 0
    SelectPoint1 = Xy2
    SelectPoint2 = AcadDoc.Utility.PolarPoint(Xy2, Ang - Pi, SearchDist)
    Set LineObj = AcadDoc.ModelSpace.AddLine(SelectPoint1, SelectPoint2)
    LineObj.Color = acBlue
    LineObj.Update
    
    DhCount = 0
    HaveDh = 0
    qq = 13
    For Each objAcadEntity In objAcadSelectionSet
        If objAcadEntity.Layer = "側溝測線" Then
            Set LwObj2 = objAcadEntity
            varReturn4 = LineObj.IntersectWith(LwObj2, acExtendOtherEntity)
            varReturn5 = LineObj.IntersectWith(LwObj2, acExtendNone)
            If UBound(varReturn4) >= 2 And LwObj2.Color = acRed Then
                HaveDh = 1
                DhCount = DhCount + 1
                m = m + 4
                StartP(0) = LwObj2.Coordinates(0)
                StartP(1) = LwObj2.Coordinates(1)
                EndP(0) = LwObj2.Coordinates(UBound(LwObj2.Coordinates) - 1)
                EndP(1) = LwObj2.Coordinates(UBound(LwObj2.Coordinates))
                
                Interpolation1 = Hdist(StartP(0), StartP(1), varReturn4(0), varReturn4(1))
                Interpolation2 = Hdist(varReturn4(0), varReturn4(1), EndP(0), EndP(1))
                Interpolation = Interpolation1 + Interpolation2
                LwObj2.Color = acMagenta
                LwObj2.Update
                r = 2
                
                Do While M_Sheet(7).Cells(r, 1).Value <> 0
                    If Abs(StartP(0) - M_Sheet(7).Cells(r, 3).Value) < 1.5 And Abs(StartP(1) - M_Sheet(7).Cells(r, 4).Value) < 1.5 And Abs(EndP(0) - M_Sheet(7).Cells(r, 6).Value) < 1.5 And Abs(EndP(1) - M_Sheet(7).Cells(r, 7).Value) < 1.5 Then
                        M_Sheet(2).Cells(i, qq).Value = M_Sheet(7).Cells(r, 2).Value
                        qq = qq + 1
                        TransectsDist = Hdist(SelectPoint2(0), SelectPoint2(1), varReturn4(0), varReturn4(1))
                        If varReturn5 >= 2 Then
                            AvgEL = Round((M_Sheet(7).Cells(r, 5).Value * Interpolation2 / Interpolation + M_Sheet(7).Cells(r, 8).Value * Interpolation1 / Interpolation), 2)
                        Else
                            If Interpolation1 < Interpolation2 Then
                                AvgEL = Round(M_Sheet(7).Cells(r, 5).Value, 2)
                            Else
                                AvgEL = Round(M_Sheet(7).Cells(r, 8).Value, 2)
                            End If
                        End If
                        AvgDepth = Round((M_Sheet(7).Cells(r, 9).Value / 2 + M_Sheet(7).Cells(r, 10).Value / 2) / 100, 2)
                        AvgWidth = Round((M_Sheet(7).Cells(r, 11).Value / 2 + M_Sheet(7).Cells(r, 12).Value / 2) / 100, 2)
                        M_Sheet(9).Cells(q, 3).Value = "Dh"
                        M_Sheet(9).Cells(q, 4).Value = Round(TransectsDist, 2)
                        M_Sheet(9).Cells(q, 5).Value = Round(AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 6).Value = Round(AvgEL, 2)
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2 + 0.01, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
                        M_Sheet(9).Cells(q, 3).Value = "Dh"
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2 - 0.01, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
                        M_Sheet(9).Cells(q, 3).Value = "Dh"
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL
                        M_Sheet(9).Cells(q, 3).Value = "Dh"
                        q = q + 1
                    ElseIf Abs(StartP(0) - M_Sheet(7).Cells(r, 6).Value) < 1.5 And Abs(StartP(1) - M_Sheet(7).Cells(r, 7).Value) < 1.5 And Abs(EndP(0) - M_Sheet(7).Cells(r, 3).Value) < 1.5 And Abs(EndP(1) - M_Sheet(7).Cells(r, 4).Value) < 1.5 Then
                        M_Sheet(2).Cells(i, qq).Value = M_Sheet(7).Cells(r, 2).Value
                        qq = qq + 1
                        TransectsDist = Hdist(SelectPoint2(0), SelectPoint2(1), varReturn4(0), varReturn4(1))
                        If varReturn5 >= 2 Then
                            AvgEL = Round((M_Sheet(7).Cells(r, 5).Value * Interpolation1 / Interpolation + M_Sheet(7).Cells(r, 8).Value * Interpolation2 / Interpolation), 2)
                        Else
                            If Interpolation1 < Interpolation2 Then
                                AvgEL = Round(M_Sheet(7).Cells(r, 8).Value, 2)
                            Else
                                AvgEL = Round(M_Sheet(7).Cells(r, 5).Value, 2)
                            End If
                        End If
                        AvgDepth = Round((M_Sheet(7).Cells(r, 9).Value / 2 + M_Sheet(7).Cells(r, 10).Value / 2) / 100, 2)
                        AvgWidth = Round((M_Sheet(7).Cells(r, 11).Value / 2 + M_Sheet(7).Cells(r, 12).Value / 2) / 100, 2)
                        M_Sheet(9).Cells(q, 3).Value = "Dh"
                        M_Sheet(9).Cells(q, 4).Value = Round(TransectsDist, 2)
                        M_Sheet(9).Cells(q, 5).Value = Round(AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 6).Value = Round(AvgEL, 2)
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2 + 0.01, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
                        M_Sheet(9).Cells(q, 3).Value = "Dh"
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2 - 0.01, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
                        M_Sheet(9).Cells(q, 3).Value = "Dh"
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL
                        M_Sheet(9).Cells(q, 3).Value = "Dh"
                        q = q + 1
                    End If
                    r = r + 1
                Loop
                LwObj2.Color = acByLayer
                LwObj2.Update
            End If
        End If
    Next

    
    FixDist1 = 0
'    For Each objAcadEntity In objAcadSelectionSet
'        If HaveDh = 0 Then
'            If objAcadEntity.Layer = "側溝測線" Then
'                Set LwObj2 = objAcadEntity
'                varReturn4 = LineObj.IntersectWith(LwObj2, acExtendOtherEntity)
'                varReturn5 = ObjRoadLine.IntersectWith(LwObj2, acExtendOtherEntity)
'                If UBound(varReturn4) >= 2 And LwObj2.Color = acRed And UBound(varReturn5) < 2 Then
'
'                    StartP(0) = LwObj2.Coordinates(0)
'                    StartP(1) = LwObj2.Coordinates(1)
'                    EndP(0) = LwObj2.Coordinates(UBound(LwObj2.Coordinates) - 1)
'                    EndP(1) = LwObj2.Coordinates(UBound(LwObj2.Coordinates))
'                    Interpolation1 = Hdist(StartP(0), StartP(1), varReturn4(0), varReturn4(1))
'                    Interpolation2 = Hdist(varReturn4(0), varReturn4(1), EndP(0), EndP(1))
'                    Interpolation = Interpolation1 + Interpolation2
'                    r = 2
'                    Do While M_Sheet(7).Cells(r, 1).Value <> 0
'                        If Abs(StartP(0) - M_Sheet(7).Cells(r, 3).Value) < 1.5 And Abs(StartP(1) - M_Sheet(7).Cells(r, 4).Value) < 1.5 And Abs(EndP(0) - M_Sheet(7).Cells(r, 6).Value) < 1.5 And Abs(EndP(1) - M_Sheet(7).Cells(r, 7).Value) < 1.5 Then
'
'                            TransectsDist = Hdist(SelectPoint2(0), SelectPoint2(1), varReturn4(0), varReturn4(1))
'                            If Interpolation1 < Interpolation2 Then
'                                AvgEL = Round(M_Sheet(7).Cells(r, 5).Value, 2)
'                            Else
'                                AvgEL = Round(M_Sheet(7).Cells(r, 8).Value, 2)
'                            End If
'                            AvgDepth = Round((M_Sheet(7).Cells(r, 9).Value / 2 + M_Sheet(7).Cells(r, 10).Value / 2) / 100, 2)
'                            AvgWidth = Round((M_Sheet(7).Cells(r, 11).Value / 2 + M_Sheet(7).Cells(r, 12).Value / 2) / 100, 2)
'                            If TransectsDist > FixDist1 / 2 - AvgWidth And TransectsDist < FixDist1 + AvgWidth / 2 Then
'                                Exit Do
'                            Else
'                                LwObj2.Color = acMagenta
'                                LwObj2.Update
'                                FixDist1 = TransectsDist
'                                DhCount = DhCount + 1
'                                m = m + 4
'                            End If
'                            M_Sheet(2).Cells(i, qq).Value = M_Sheet(7).Cells(r, 2).Value
'                            qq = qq + 1
'                            M_Sheet(9).Cells(q, 4).Value = Round(TransectsDist, 2)
'                            M_Sheet(9).Cells(q, 5).Value = Round(AvgWidth / 2, 2)
'                            M_Sheet(9).Cells(q, 6).Value = Round(AvgEL, 2)
'                            M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2, 2)
'                            M_Sheet(9).Cells(q, 2).Value = AvgEL
'                            q = q + 1
'                            M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2 + 0.01, 2)
'                            M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
'                            q = q + 1
'                            M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2 - 0.01, 2)
'                            M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
'                            q = q + 1
'                            M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2, 2)
'                            M_Sheet(9).Cells(q, 2).Value = AvgEL
'                            q = q + 1
'                        ElseIf Abs(StartP(0) - M_Sheet(7).Cells(r, 6).Value) < 1.5 And Abs(StartP(1) - M_Sheet(7).Cells(r, 7).Value) < 1.5 And Abs(EndP(0) - M_Sheet(7).Cells(r, 3).Value) < 1.5 And Abs(EndP(1) - M_Sheet(7).Cells(r, 4).Value) < 1.5 Then
'
'                            TransectsDist = Hdist(SelectPoint2(0), SelectPoint2(1), varReturn4(0), varReturn4(1))
'                            If Interpolation1 < Interpolation2 Then
'                                AvgEL = Round(M_Sheet(7).Cells(r, 8).Value, 2)
'                            Else
'                                AvgEL = Round(M_Sheet(7).Cells(r, 5).Value, 2)
'                            End If
'                            AvgDepth = Round((M_Sheet(7).Cells(r, 9).Value / 2 + M_Sheet(7).Cells(r, 10).Value / 2) / 100, 2)
'                            AvgWidth = Round((M_Sheet(7).Cells(r, 11).Value / 2 + M_Sheet(7).Cells(r, 12).Value / 2) / 100, 2)
'                            If TransectsDist > FixDist1 / 2 - AvgWidth And TransectsDist < FixDist1 + AvgWidth / 2 Then
'                                Exit Do
'                            Else
'                                LwObj2.Color = acMagenta
'                                LwObj2.Update
'                                FixDist1 = TransectsDist
'                                DhCount = DhCount + 1
'                                m = m + 4
'                            End If
'                            M_Sheet(2).Cells(i, qq).Value = M_Sheet(7).Cells(r, 2).Value
'                            qq = qq + 1
'                            M_Sheet(9).Cells(q, 4).Value = Round(TransectsDist, 2)
'                            M_Sheet(9).Cells(q, 5).Value = Round(AvgWidth / 2, 2)
'                            M_Sheet(9).Cells(q, 6).Value = Round(AvgEL, 2)
'                            M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2, 2)
'                            M_Sheet(9).Cells(q, 2).Value = AvgEL
'                            q = q + 1
'                            M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2 + 0.01, 2)
'                            M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
'                            q = q + 1
'                            M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2 - 0.01, 2)
'                            M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
'                            q = q + 1
'                            M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2, 2)
'                            M_Sheet(9).Cells(q, 2).Value = AvgEL
'                            q = q + 1
'                        End If
'                        r = r + 1
'                    Loop
'
'                    LwObj2.Color = acByLayer
'                    LwObj2.Update
'                End If
'            End If
'        End If
'    Next
    LineObj.Delete
    
    j = 3
    Dh = q - 1
    Do While M_Sheet(8).Cells(j, 1).Value <> ""
        If Abs(Xy3(0) - M_Sheet(8).Cells(j, 2).Value) < 1 And Abs(Xy3(1) - M_Sheet(8).Cells(j, 3).Value) < 1 Then
            k = 2
            'Do While M_Sheet(8).Cells(j, k).Value <> ""
            For k = 2 To 22
                If M_Sheet(8).Cells(j, k).Value <> "" Then
                    TransectsDist = Hdist(SelectPoint2(0), SelectPoint2(1), M_Sheet(8).Cells(j, k).Value, M_Sheet(8).Cells(j, k + 1).Value)
                    M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist, 2)
                    M_Sheet(9).Cells(q, 2).Value = Round(M_Sheet(8).Cells(j, k + 2).Value, 2)
                    M_Sheet(9).Cells(q, 3).Value = "Road"
                    For Dh0 = 2 To Dh
                        If Abs(M_Sheet(9).Cells(q, 1).Value - M_Sheet(9).Cells(Dh0, 4).Value) <= Abs(M_Sheet(9).Cells(Dh0, 5).Value) Then
                            If M_Sheet(9).Cells(q, 1).Value - M_Sheet(9).Cells(Dh0, 4).Value <= 0 Then
                                M_Sheet(9).Cells(Dh0, 1).Value = M_Sheet(9).Cells(Dh0, 1).Value + M_Sheet(9).Cells(Dh0, 5).Value - Abs(M_Sheet(9).Cells(q, 1).Value - M_Sheet(9).Cells(Dh0, 4).Value) + 0.01
                                M_Sheet(9).Cells(Dh0 + 1, 1).Value = M_Sheet(9).Cells(Dh0 + 1, 1).Value + M_Sheet(9).Cells(Dh0, 5).Value - Abs(M_Sheet(9).Cells(q, 1).Value - M_Sheet(9).Cells(Dh0, 4).Value) + 0.01
                                M_Sheet(9).Cells(Dh0 + 2, 1).Value = M_Sheet(9).Cells(Dh0 + 2, 1).Value + M_Sheet(9).Cells(Dh0, 5).Value - Abs(M_Sheet(9).Cells(q, 1).Value - M_Sheet(9).Cells(Dh0, 4).Value) + 0.01
                                M_Sheet(9).Cells(Dh0 + 3, 1).Value = M_Sheet(9).Cells(Dh0 + 3, 1).Value + M_Sheet(9).Cells(Dh0, 5).Value - Abs(M_Sheet(9).Cells(q, 1).Value - M_Sheet(9).Cells(Dh0, 4).Value) + 0.01
                           
                            ElseIf M_Sheet(9).Cells(q, 1).Value - M_Sheet(9).Cells(Dh0, 4).Value > 0 Then
                                M_Sheet(9).Cells(Dh0, 1).Value = M_Sheet(9).Cells(Dh0, 1).Value + M_Sheet(9).Cells(Dh0, 5).Value + Abs(M_Sheet(9).Cells(q, 1).Value - M_Sheet(9).Cells(Dh0, 4).Value) + 0.01
                                M_Sheet(9).Cells(Dh0 + 1, 1).Value = M_Sheet(9).Cells(Dh0 + 1, 1).Value + M_Sheet(9).Cells(Dh0, 5).Value + Abs(M_Sheet(9).Cells(q, 1).Value - M_Sheet(9).Cells(Dh0, 4).Value) + 0.01
                                M_Sheet(9).Cells(Dh0 + 2, 1).Value = M_Sheet(9).Cells(Dh0 + 2, 1).Value + M_Sheet(9).Cells(Dh0, 5).Value + Abs(M_Sheet(9).Cells(q, 1).Value - M_Sheet(9).Cells(Dh0, 4).Value) + 0.01
                                M_Sheet(9).Cells(Dh0 + 3, 1).Value = M_Sheet(9).Cells(Dh0 + 3, 1).Value + M_Sheet(9).Cells(Dh0, 5).Value + Abs(M_Sheet(9).Cells(q, 1).Value - M_Sheet(9).Cells(Dh0, 4).Value) + 0.01
                            
                            End If
                        End If
                    Next
                    q = q + 1
                    k = k + 2
                End If
            'Loop
            Next
        End If
        j = j + 1
    Loop
    
    Dh1 = q - 1
    Dh2 = q
    SelectPoint3 = SelectPoint2
    Ang = AzToAcadAngle(Pol(LwObj11.Coordinates(UBound(LwObj11.Coordinates)), LwObj11.Coordinates(UBound(LwObj11.Coordinates) - 1), Xy3(1), Xy3(0)))
    '抓終點側側溝--------------------------------------
    Set objAcadSelectionSet = AcadDoc.SelectionSets.Add("DhSelectionSetFilter")
    'MsgBox (objAcadSelectionSet.Count)
    Xy2(0) = Xy3(0)
    Xy2(1) = Xy3(1)
    Xy2(2) = 0
    SelectPoint1 = Xy2
    SelectPoint2 = AcadDoc.Utility.PolarPoint(Xy2, Ang - Pi, SearchDist)
    Set LineObj = AcadDoc.ModelSpace.AddLine(SelectPoint1, SelectPoint2)
    LineObj.Color = acBlue
    LineObj.Update
    
    HaveDh = 0
    For Each objAcadEntity In objAcadSelectionSet
        If objAcadEntity.Layer = "側溝測線" Then
            Set LwObj2 = objAcadEntity
            varReturn4 = LineObj.IntersectWith(LwObj2, acExtendOtherEntity)
            varReturn5 = LineObj.IntersectWith(LwObj2, acExtendNone)
            If UBound(varReturn4) >= 2 And LwObj2.Color = acRed Then
                HaveDh = 1
                DhCount = DhCount + 1
                m = m + 4
                StartP(0) = LwObj2.Coordinates(0)
                StartP(1) = LwObj2.Coordinates(1)
                EndP(0) = LwObj2.Coordinates(UBound(LwObj2.Coordinates) - 1)
                EndP(1) = LwObj2.Coordinates(UBound(LwObj2.Coordinates))
                Interpolation1 = Hdist(StartP(0), StartP(1), varReturn4(0), varReturn4(1))
                Interpolation2 = Hdist(varReturn4(0), varReturn4(1), EndP(0), EndP(1))
                Interpolation = Interpolation1 + Interpolation2
                LwObj2.Color = acMagenta
                LwObj2.Update
                r = 2
                Do While M_Sheet(7).Cells(r, 1).Value <> 0
                    If Abs(StartP(0) - M_Sheet(7).Cells(r, 3).Value) < 1.5 And Abs(StartP(1) - M_Sheet(7).Cells(r, 4).Value) < 1.5 And Abs(EndP(0) - M_Sheet(7).Cells(r, 6).Value) < 1.5 And Abs(EndP(1) - M_Sheet(7).Cells(r, 7).Value) < 1.5 Then
                        M_Sheet(2).Cells(i, qq).Value = M_Sheet(7).Cells(r, 2).Value
                        qq = qq + 1
                        TransectsDist = Hdist(SelectPoint3(0), SelectPoint3(1), varReturn4(0), varReturn4(1))
                        If varReturn5 >= 2 Then
                            AvgEL = Round((M_Sheet(7).Cells(r, 5).Value * Interpolation2 / Interpolation + M_Sheet(7).Cells(r, 8).Value * Interpolation1 / Interpolation), 2)
                        Else
                            If Interpolation1 < Interpolation2 Then
                                AvgEL = Round(M_Sheet(7).Cells(r, 5).Value, 2)
                            Else
                                AvgEL = Round(M_Sheet(7).Cells(r, 8).Value, 2)
                            End If
                        End If
                        AvgDepth = Round((M_Sheet(7).Cells(r, 9).Value / 2 + M_Sheet(7).Cells(r, 10).Value / 2) / 100, 2)
                        AvgWidth = Round((M_Sheet(7).Cells(r, 11).Value / 2 + M_Sheet(7).Cells(r, 12).Value / 2) / 100, 2)
                        M_Sheet(9).Cells(q, 3).Value = "Dh"
                        M_Sheet(9).Cells(q, 4).Value = Round(TransectsDist, 2)
                        M_Sheet(9).Cells(q, 5).Value = Round(AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 6).Value = Round(AvgEL, 2)
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2 + 0.01, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
                        M_Sheet(9).Cells(q, 3).Value = "Dh"
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2 - 0.01, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
                        M_Sheet(9).Cells(q, 3).Value = "Dh"
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL
                        M_Sheet(9).Cells(q, 3).Value = "Dh"
                        q = q + 1
                    ElseIf Abs(StartP(0) - M_Sheet(7).Cells(r, 6).Value) < 1.5 And Abs(StartP(1) - M_Sheet(7).Cells(r, 7).Value) < 1.5 And Abs(EndP(0) - M_Sheet(7).Cells(r, 3).Value) < 1.5 And Abs(EndP(1) - M_Sheet(7).Cells(r, 4).Value) < 1.5 Then
                        M_Sheet(2).Cells(i, qq).Value = M_Sheet(7).Cells(r, 2).Value
                        qq = qq + 1
                        TransectsDist = Hdist(SelectPoint3(0), SelectPoint3(1), varReturn4(0), varReturn4(1))
                        If varReturn5 >= 2 Then
                            AvgEL = Round((M_Sheet(7).Cells(r, 5).Value * Interpolation1 / Interpolation + M_Sheet(7).Cells(r, 8).Value * Interpolation2 / Interpolation), 2)
                        Else
                            If Interpolation1 < Interpolation2 Then
                                AvgEL = Round(M_Sheet(7).Cells(r, 8).Value, 2)
                            Else
                                AvgEL = Round(M_Sheet(7).Cells(r, 5).Value, 2)
                            End If
                        End If
                        AvgDepth = Round((M_Sheet(7).Cells(r, 9).Value / 2 + M_Sheet(7).Cells(r, 10).Value / 2) / 100, 2)
                        AvgWidth = Round((M_Sheet(7).Cells(r, 11).Value / 2 + M_Sheet(7).Cells(r, 12).Value / 2) / 100, 2)
                        M_Sheet(9).Cells(q, 3).Value = "Dh"
                        M_Sheet(9).Cells(q, 4).Value = Round(TransectsDist, 2)
                        M_Sheet(9).Cells(q, 5).Value = Round(AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 6).Value = Round(AvgEL, 2)
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2 + 0.01, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
                        M_Sheet(9).Cells(q, 3).Value = "Dh"
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2 - 0.01, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
                        M_Sheet(9).Cells(q, 3).Value = "Dh"
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL
                        M_Sheet(9).Cells(q, 3).Value = "Dh"
                        q = q + 1
                    End If
                    r = r + 1
                Loop
                LwObj2.Color = acByLayer
                LwObj2.Update
            End If
        End If
    Next
    
    FixDist1 = 0
'    For Each objAcadEntity In objAcadSelectionSet
'        If HaveDh = 0 Then
'            If objAcadEntity.Layer = "側溝測線" Then
'                Set LwObj2 = objAcadEntity
'                varReturn4 = LineObj.IntersectWith(LwObj2, acExtendOtherEntity)
'                varReturn5 = ObjRoadLine.IntersectWith(LwObj2, acExtendOtherEntity)
'                If UBound(varReturn4) >= 2 And LwObj2.Color = acRed And UBound(varReturn5) < 2 Then
'                    StartP(0) = LwObj2.Coordinates(0)
'                    StartP(1) = LwObj2.Coordinates(1)
'                    EndP(0) = LwObj2.Coordinates(UBound(LwObj2.Coordinates) - 1)
'                    EndP(1) = LwObj2.Coordinates(UBound(LwObj2.Coordinates))
'                    Interpolation1 = Hdist(StartP(0), StartP(1), varReturn4(0), varReturn4(1))
'                    Interpolation2 = Hdist(varReturn4(0), varReturn4(1), EndP(0), EndP(1))
'                    Interpolation = Interpolation1 + Interpolation2
'                    r = 2
'                    Do While M_Sheet(7).Cells(r, 1).Value <> 0
'                        If Abs(StartP(0) - M_Sheet(7).Cells(r, 3).Value) < 1.5 And Abs(StartP(1) - M_Sheet(7).Cells(r, 4).Value) < 1.5 And Abs(EndP(0) - M_Sheet(7).Cells(r, 6).Value) < 1.5 And Abs(EndP(1) - M_Sheet(7).Cells(r, 7).Value) < 1.5 Then
'                            TransectsDist = Hdist(SelectPoint3(0), SelectPoint3(1), varReturn4(0), varReturn4(1))
'                            If Interpolation1 < Interpolation2 Then
'                                AvgEL = Round(M_Sheet(7).Cells(r, 5).Value, 2)
'                            Else
'                                AvgEL = Round(M_Sheet(7).Cells(r, 8).Value, 2)
'                            End If
'                            AvgDepth = Round((M_Sheet(7).Cells(r, 9).Value / 2 + M_Sheet(7).Cells(r, 10).Value / 2) / 100, 2)
'                            AvgWidth = Round((M_Sheet(7).Cells(r, 11).Value / 2 + M_Sheet(7).Cells(r, 12).Value / 2) / 100, 2)
'                            If TransectsDist > FixDist1 / 2 - AvgWidth And TransectsDist < FixDist1 + AvgWidth / 2 Then
'                                Exit Do
'                            Else
'                                LwObj2.Color = acMagenta
'                                LwObj2.Update
'                                FixDist1 = TransectsDist
'                                DhCount = DhCount + 1
'                                m = m + 4
'                            End If
'                            M_Sheet(2).Cells(i, qq).Value = M_Sheet(7).Cells(r, 2).Value
'                            qq = qq + 1
'                            M_Sheet(9).Cells(q, 4).Value = Round(TransectsDist, 2)
'                            M_Sheet(9).Cells(q, 5).Value = Round(AvgWidth / 2, 2)
'                            M_Sheet(9).Cells(q, 6).Value = Round(AvgEL, 2)
'                            M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2, 2)
'                            M_Sheet(9).Cells(q, 2).Value = AvgEL
'                            q = q + 1
'                            M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2 + 0.01, 2)
'                            M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
'                            q = q + 1
'                            M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2 - 0.01, 2)
'                            M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
'                            q = q + 1
'                            M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2, 2)
'                            M_Sheet(9).Cells(q, 2).Value = AvgEL
'                            q = q + 1
'                        ElseIf Abs(StartP(0) - M_Sheet(7).Cells(r, 6).Value) < 1.5 And Abs(StartP(1) - M_Sheet(7).Cells(r, 7).Value) < 1.5 And Abs(EndP(0) - M_Sheet(7).Cells(r, 3).Value) < 1.5 And Abs(EndP(1) - M_Sheet(7).Cells(r, 4).Value) < 1.5 Then
'
'                            TransectsDist = Hdist(SelectPoint3(0), SelectPoint3(1), varReturn4(0), varReturn4(1))
'                            If Interpolation1 < Interpolation2 Then
'                                AvgEL = Round(M_Sheet(7).Cells(r, 8).Value, 2)
'                            Else
'                                AvgEL = Round(M_Sheet(7).Cells(r, 5).Value, 2)
'                            End If
'                            AvgDepth = Round((M_Sheet(7).Cells(r, 9).Value / 2 + M_Sheet(7).Cells(r, 10).Value / 2) / 100, 2)
'                            AvgWidth = Round((M_Sheet(7).Cells(r, 11).Value / 2 + M_Sheet(7).Cells(r, 12).Value / 2) / 100, 2)
'                            If TransectsDist > FixDist1 / 2 - AvgWidth And TransectsDist < FixDist1 + AvgWidth / 2 Then
'                                Exit Do
'                            Else
'                                LwObj2.Color = acMagenta
'                                LwObj2.Update
'                                FixDist1 = TransectsDist
'                                DhCount = DhCount + 1
'                                m = m + 4
'                            End If
'                            M_Sheet(2).Cells(i, qq).Value = M_Sheet(7).Cells(r, 2).Value
'                            qq = qq + 1
'                            M_Sheet(9).Cells(q, 4).Value = Round(TransectsDist, 2)
'                            M_Sheet(9).Cells(q, 5).Value = Round(AvgWidth / 2, 2)
'                            M_Sheet(9).Cells(q, 6).Value = Round(AvgEL, 2)
'                            M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2, 2)
'                            M_Sheet(9).Cells(q, 2).Value = AvgEL
'                            q = q + 1
'                            M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2 + 0.01, 2)
'                            M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
'                            q = q + 1
'                            M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2 - 0.01, 2)
'                            M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
'                            q = q + 1
'                            M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2, 2)
'                            M_Sheet(9).Cells(q, 2).Value = AvgEL
'                            q = q + 1
'                        End If
'                        r = r + 1
'                    Loop
'
'                    LwObj2.Color = acByLayer
'                    LwObj2.Update
'                End If
'            End If
'        End If
'    Next
    LineObj.Delete
    
'    For Each objAcadEntity In objAcadSelectionSet
'        Set LwObj1 = objAcadEntity
'        LwObj1.Color = acByLayer
'        LwObj1.Update
'    Next
 
For Dh0 = Dh + 1 To Dh1
    For Dh00 = Dh2 To q
        If Abs(M_Sheet(9).Cells(Dh0, 1).Value - M_Sheet(9).Cells(Dh00, 4).Value) <= Abs(M_Sheet(9).Cells(Dh00, 5).Value) Then
            If M_Sheet(9).Cells(Dh0, 1).Value - M_Sheet(9).Cells(Dh00, 4).Value <= 0 Then
                M_Sheet(9).Cells(Dh00, 1).Value = M_Sheet(9).Cells(Dh00, 1).Value - M_Sheet(9).Cells(Dh00, 5).Value - Abs(M_Sheet(9).Cells(Dh0, 1).Value - M_Sheet(9).Cells(Dh00, 4).Value) - 0.01
                M_Sheet(9).Cells(Dh00 + 1, 1).Value = M_Sheet(9).Cells(Dh00 + 1, 1).Value - M_Sheet(9).Cells(Dh00, 5).Value - Abs(M_Sheet(9).Cells(Dh0, 1).Value - M_Sheet(9).Cells(Dh00, 4).Value) - 0.01
                M_Sheet(9).Cells(Dh00 + 2, 1).Value = M_Sheet(9).Cells(Dh00 + 2, 1).Value - M_Sheet(9).Cells(Dh00, 5).Value - Abs(M_Sheet(9).Cells(Dh0, 1).Value - M_Sheet(9).Cells(Dh00, 4).Value) - 0.01
                M_Sheet(9).Cells(Dh00 + 3, 1).Value = M_Sheet(9).Cells(Dh00 + 3, 1).Value - M_Sheet(9).Cells(Dh00, 5).Value - Abs(M_Sheet(9).Cells(Dh0, 1).Value - M_Sheet(9).Cells(Dh00, 4).Value) - 0.01
           
            ElseIf M_Sheet(9).Cells(Dh0, 1).Value - M_Sheet(9).Cells(Dh00, 4).Value > 0 Then
                M_Sheet(9).Cells(Dh00, 1).Value = M_Sheet(9).Cells(Dh00, 1).Value - M_Sheet(9).Cells(Dh00, 5).Value + Abs(M_Sheet(9).Cells(Dh0, 1).Value - M_Sheet(9).Cells(Dh00, 4).Value) - 0.01
                M_Sheet(9).Cells(Dh00 + 1, 1).Value = M_Sheet(9).Cells(Dh00 + 1, 1).Value - M_Sheet(9).Cells(Dh00, 5).Value + Abs(M_Sheet(9).Cells(Dh0, 1).Value - M_Sheet(9).Cells(Dh00, 4).Value) - 0.01
                M_Sheet(9).Cells(Dh00 + 2, 1).Value = M_Sheet(9).Cells(Dh00 + 2, 1).Value - M_Sheet(9).Cells(Dh00, 5).Value + Abs(M_Sheet(9).Cells(Dh0, 1).Value - M_Sheet(9).Cells(Dh00, 4).Value) - 0.01
                M_Sheet(9).Cells(Dh00 + 3, 1).Value = M_Sheet(9).Cells(Dh00 + 3, 1).Value - M_Sheet(9).Cells(Dh00, 5).Value + Abs(M_Sheet(9).Cells(Dh0, 1).Value - M_Sheet(9).Cells(Dh00, 4).Value) - 0.01
            
            End If
        End If
    Next
Next
    
    ReDim Transects(0 To n + m)
    M_Sheet(9).Activate
    Columns("A:A").Select
    M_Sheet(9).Sort.SortFields.Clear
    M_Sheet(9).Sort.SortFields.Add Key:=Range("A2:A" & q - 1) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet2").Sort
        .SetRange Range("A2:C" & q - 1)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    FixDist = M_Sheet(9).Cells(u, 1).Value
    If FixDist > 10 Then
        Do While M_Sheet(9).Cells(u, 1).Value <> ""
            M_Sheet(9).Cells(u, 1).Value = M_Sheet(9).Cells(u, 1).Value - FixDist + 10
            u = u + 1
        Loop
    End If

    u = 2
    FixI = w
    M_Sheet(4).Activate
    M_Sheet(4).Cells(w, 1).Value = "NC"
    M_Sheet(4).Cells(w, 2).Value = "0.015"
    M_Sheet(4).Cells(w, 3).Value = "0.015"
    M_Sheet(4).Cells(w, 4).Value = "0.011"
    w = w + 1
    M_Sheet(4).Cells(w, 1).Value = "X1"
    M_Sheet(4).Cells(w, 2).Value = M_Sheet(6).Cells(i, 1).Value
    M_Sheet(4).Cells(w, 4).Value = (n + 1) / 2 + m
    M_Sheet(4).Cells(w, 5).Value = M_Sheet(9).Cells(u, 1).Value
    Do While M_Sheet(9).Cells(u, 1).Value <> ""
        M_Sheet(4).Cells(w, 6).Value = M_Sheet(9).Cells(u, 1).Value
        u = u + 1
    Loop
    M_Sheet(9).Cells(u, 1).Value = M_Sheet(9).Cells(u - 1, 1).Value + 10
    M_Sheet(9).Cells(u, 2).Value = M_Sheet(9).Cells(u - 1, 2).Value + 3
    'M_Sheet(4).Cells(w, 5).Value = M_Sheet(9).Cells(u, 1).Value
    u = 2
    M_Sheet(4).Cells(w, 7).Value = "0"
    M_Sheet(4).Cells(w, 8).Value = "0"
    M_Sheet(4).Cells(w, 9).Value = "0"
    M_Sheet(4).Cells(w, 10).Value = "1"
    M_Sheet(4).Cells(w, 11).Value = "0"
    w = w + 1
    M_Sheet(4).Cells(w, 1).Value = "GR"
    M_Sheet(4).Cells(w, 2).Value = M_Sheet(9).Cells(u, 2).Value + 3
    M_Sheet(4).Cells(w, 3).Value = "0"
    'w = w - 1
    ww = 4
    CircleNum = 1
    LineNum = 1
    Xy222(0) = 0
    Xy222(1) = 0
    Xy222(2) = 0
    Do While M_Sheet(9).Cells(u, 1) <> ""
        Xy22(2) = 0
        M_Sheet(4).Cells(w, ww).Value = M_Sheet(9).Cells(u, 2).Value
        Xy22(1) = M_Sheet(9).Cells(u, 2).Value
        ww = ww + 1
        M_Sheet(4).Cells(w, ww).Value = M_Sheet(9).Cells(u, 1).Value
        Xy22(0) = M_Sheet(9).Cells(u, 1).Value
        ww = ww + 1
        u = u + 1
        Set ObjCircle(CircleNum) = AcadDoc1.ModelSpace.AddCircle(Xy22, 0.1)
        ObjCircle(CircleNum).Update
        CircleNum = CircleNum + 1
        If Xy222(0) <> 0 Or Xy222(1) <> 0 Then
            Set ObjLine(LineNum) = AcadDoc1.ModelSpace.AddLine(Xy22, Xy222)
            ObjLine(LineNum).Update
            LineNum = LineNum + 1
        End If
        Xy222(0) = Xy22(0)
        Xy222(1) = Xy22(1)
        Xy222(2) = 0
        If ww = 12 And M_Sheet(9).Cells(u, 1).Value <> "" Then
            w = w + 1
            ww = 1
            M_Sheet(4).Cells(w, ww).Value = "GR"
            ww = ww + 1
            M_Sheet(4).Cells(w, ww).Value = M_Sheet(9).Cells(u, 2).Value
            Xy22(1) = M_Sheet(9).Cells(u, 2).Value
            ww = ww + 1
            M_Sheet(4).Cells(w, ww).Value = M_Sheet(9).Cells(u, 1).Value
            Xy22(0) = M_Sheet(9).Cells(u, 1).Value
            ww = ww + 1
            u = u + 1
            Set ObjCircle(CircleNum) = AcadDoc1.ModelSpace.AddCircle(Xy22, 0.1)
            ObjCircle(CircleNum).Update
            CircleNum = CircleNum + 1
            If Xy222(0) <> 0 Or Xy222(1) <> 0 Then
                Set ObjLine(LineNum) = AcadDoc1.ModelSpace.AddLine(Xy22, Xy222)
                ObjLine(LineNum).Update
                LineNum = LineNum + 1
            End If
            Xy222(0) = Xy22(0)
            Xy222(1) = Xy22(1)
            Xy222(2) = 0
        End If
    Loop
    
    u = 2
    SendKeys "%{TAB}"

    FixOrNot = InputBox("Fix?")
    If FixOrNot = -1 Then
    Stop
    ElseIf FixOrNot <> "" Then
    w = FixI
    For qqq = 1 To CircleNum - 1
        ObjCircle(qqq).Delete
            'ObjCircle(q).Update
    Next
    For qqq = 1 To LineNum - 1
        ObjLine(qqq).Delete
            'ObjLine(q).Update
    Next
    Do While M_Sheet(9).Cells(u, 2).Value <> ""
        FixEL1 = -1
        FixEL2 = -1
        If M_Sheet(9).Cells(u, 3).Value = "Dh" Then
            FixEL1 = M_Sheet(9).Cells(u, 1).Value - M_Sheet(9).Cells(u - 1, 1).Value
            If M_Sheet(9).Cells(u - 1, 1).Value = "" Or FixEL1 = -1 Then
                FixEL1 = 100
            End If
            FixEL2 = M_Sheet(9).Cells(u + 4, 1).Value - M_Sheet(9).Cells(u + 3, 1).Value
            If M_Sheet(9).Cells(u + 4, 1).Value = "" Or FixEL2 = -1 Then
                FixEL2 = 100
            End If
            If FixEL1 < FixEL2 And M_Sheet(9).Cells(u - 1, 3).Value <> "Dh" Then
                M_Sheet(9).Cells(u, 2).Value = M_Sheet(9).Cells(u - 1, 2).Value
                M_Sheet(9).Cells(u + 1, 2).Value = M_Sheet(9).Cells(u + 1, 2).Value + (M_Sheet(9).Cells(u - 1, 2).Value - M_Sheet(9).Cells(u + 3, 2).Value)
                M_Sheet(9).Cells(u + 2, 2).Value = M_Sheet(9).Cells(u + 2, 2).Value + (M_Sheet(9).Cells(u - 1, 2).Value - M_Sheet(9).Cells(u + 3, 2).Value)
                M_Sheet(9).Cells(u + 3, 2).Value = M_Sheet(9).Cells(u + 3, 2).Value + (M_Sheet(9).Cells(u - 1, 2).Value - M_Sheet(9).Cells(u + 3, 2).Value)
                u = u + 4
            ElseIf FixEL1 > FixEL2 And M_Sheet(9).Cells(u + 4, 3).Value <> "Dh" Then
                M_Sheet(9).Cells(u, 2).Value = M_Sheet(9).Cells(u + 4, 2).Value
                M_Sheet(9).Cells(u + 1, 2).Value = M_Sheet(9).Cells(u + 1, 2).Value + (M_Sheet(9).Cells(u + 4, 2).Value - M_Sheet(9).Cells(u + 3, 2).Value)
                M_Sheet(9).Cells(u + 2, 2).Value = M_Sheet(9).Cells(u + 2, 2).Value + (M_Sheet(9).Cells(u + 4, 2).Value - M_Sheet(9).Cells(u + 3, 2).Value)
                M_Sheet(9).Cells(u + 3, 2).Value = M_Sheet(9).Cells(u + 3, 2).Value + (M_Sheet(9).Cells(u + 4, 2).Value - M_Sheet(9).Cells(u + 3, 2).Value)
                u = u + 4
            Else
                u = u + 1
            End If
        Else
            u = u + 1
        End If
    Loop
        u = 2
    

    M_Sheet(4).Cells(w, 1).Value = "NC"
    M_Sheet(4).Cells(w, 2).Value = "0.015"
    M_Sheet(4).Cells(w, 3).Value = "0.015"
    M_Sheet(4).Cells(w, 4).Value = "0.011"
    w = w + 1
    M_Sheet(4).Cells(w, 1).Value = "X1"
    M_Sheet(4).Cells(w, 2).Value = M_Sheet(6).Cells(i, 1).Value
    M_Sheet(4).Cells(w, 4).Value = (n + 1) / 2 + m
    M_Sheet(4).Cells(w, 5).Value = M_Sheet(9).Cells(u, 1).Value
    Do While M_Sheet(9).Cells(u, 1).Value <> ""
        M_Sheet(4).Cells(w, 6).Value = M_Sheet(9).Cells(u, 1).Value
        u = u + 1
    Loop
    'M_Sheet(4).Cells(w, 5).Value = M_Sheet(9).Cells(u, 1).Value
    u = 2
    M_Sheet(4).Cells(w, 7).Value = "0"
    M_Sheet(4).Cells(w, 8).Value = "0"
    M_Sheet(4).Cells(w, 9).Value = "0"
    M_Sheet(4).Cells(w, 10).Value = "1"
    M_Sheet(4).Cells(w, 11).Value = "0"
    w = w + 1
    M_Sheet(4).Cells(w, 1).Value = "GR"
    M_Sheet(4).Cells(w, 2).Value = M_Sheet(9).Cells(u, 2).Value + 3
    M_Sheet(4).Cells(w, 3).Value = "0"
    'w = w - 1
    ww = 4
    CircleNum = 1
    LineNum = 1
    Xy222(0) = 0
    Xy222(1) = 0
    Xy222(2) = 0
    Do While M_Sheet(9).Cells(u, 1) <> ""
        Xy22(2) = 0
        M_Sheet(4).Cells(w, ww).Value = M_Sheet(9).Cells(u, 2).Value
        Xy22(1) = M_Sheet(9).Cells(u, 2).Value
        ww = ww + 1
        M_Sheet(4).Cells(w, ww).Value = M_Sheet(9).Cells(u, 1).Value
        Xy22(0) = M_Sheet(9).Cells(u, 1).Value
        ww = ww + 1
        u = u + 1
        Set ObjCircle(CircleNum) = AcadDoc1.ModelSpace.AddCircle(Xy22, 0.1)
        ObjCircle(CircleNum).Update
        CircleNum = CircleNum + 1
        If Xy222(0) <> 0 Or Xy222(1) <> 0 Then
            Set ObjLine(LineNum) = AcadDoc1.ModelSpace.AddLine(Xy22, Xy222)
            ObjLine(LineNum).Update
            LineNum = LineNum + 1
        End If
        Xy222(0) = Xy22(0)
        Xy222(1) = Xy22(1)
        Xy222(2) = 0
        If ww = 12 And M_Sheet(9).Cells(u, 1).Value <> "" Then
            w = w + 1
            ww = 1
            M_Sheet(4).Cells(w, ww).Value = "GR"
            ww = ww + 1
            M_Sheet(4).Cells(w, ww).Value = M_Sheet(9).Cells(u, 2).Value
            Xy22(1) = M_Sheet(9).Cells(u, 2).Value
            ww = ww + 1
            M_Sheet(4).Cells(w, ww).Value = M_Sheet(9).Cells(u, 1).Value
            Xy22(0) = M_Sheet(9).Cells(u, 1).Value
            ww = ww + 1
            u = u + 1
            Set ObjCircle(CircleNum) = AcadDoc1.ModelSpace.AddCircle(Xy22, 0.1)
            ObjCircle(CircleNum).Update
            CircleNum = CircleNum + 1
            If Xy222(0) <> 0 Or Xy222(1) <> 0 Then
                Set ObjLine(LineNum) = AcadDoc1.ModelSpace.AddLine(Xy22, Xy222)
                ObjLine(LineNum).Update
                LineNum = LineNum + 1
            End If
            Xy222(0) = Xy22(0)
            Xy222(1) = Xy22(1)
            Xy222(2) = 0
        End If
    Loop
    MsgBox ("")
    End If
    w = w + 1
    M_Sheet(4).Cells(w, 1).Value = ";"
    w = w + 1
    M_Sheet(4).Cells(w, 1).Select
    
 'Stop
T:
    For qqq = 1 To CircleNum - 1
        ObjCircle(qqq).Delete
            'ObjCircle(q).Update
    Next
    For qqq = 1 To LineNum - 1
        ObjLine(qqq).Delete
            'ObjLine(q).Update
    Next
    LwObj00.Delete
    LwObj01.Delete
    LwObj02.Delete
    LwObj03.Delete
    M_Sheet(9).Activate
    Columns("A:H").Select
    Selection.Delete Shift:=xlToLeft
    M_Sheet(4).Activate
Z:
    For Each objAcadEntity In objAcadSelectionSet
        Set LwObj1 = objAcadEntity
        LwObj1.Color = acByLayer
        LwObj1.Update
    Next
    M_Sheet(2).Cells(i, 11) = DhCount
    DhCount = 0
    SendKeys "%{TAB}"

    i = i + 1
    ObjRoadLine.Delete
    'w = 40474
    'i = 6941
    
Loop

End Sub

