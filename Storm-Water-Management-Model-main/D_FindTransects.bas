Attribute VB_Name = "D_FindTransects"
Sub FindTransects()
Dim Xy2(0 To 2) As Double
Dim Xy3(0 To 3) As Double
Dim SelectPoint1 As Variant
Dim SelectPoint2 As Variant
Dim SelectPoint3 As Variant
Dim LwObj As AcadLWPolyline
Dim LwObj00 As AcadLWPolyline
Dim LwObj01 As AcadLWPolyline
Dim LwObj02 As AcadLWPolyline
Dim LwObj03 As AcadLWPolyline
Dim LwObj1 As AcadLWPolyline
Dim LwObj2 As AcadLWPolyline
Dim ZoomPoint1(0 To 2) As Double
Dim objAcadSelectionSet As AcadSelectionSet
Dim objAcadEntity As AcadEntity
Dim varReturn As Variant
Dim varReturn1 As Variant
Dim varReturn2 As Variant
Dim varReturn3 As Variant
Dim varReturn4 As Variant
Dim Transects() As Variant
Dim SearchDist As Double
Dim LineObj As AcadLine
Dim TransectsDist As Double
Dim AvgEL As Double
Dim AvgWidth As Double
Dim AvgDepth As Double
Dim StartP(0 To 1), EndP(0 To 1) As Double
Dim FixDist As Double
Dim DhCount, Dh, Dh0, Dh00, Dh1, Dh2 As Integer



i = 1
w = 1
Do While M_Sheet(6).Cells(i, 1) <> ""
    'M_Sheet(6).Cells(i, 1).Select
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
    If n = 0 Then
        GoTo Z
    End If
    
    
    q = 2
    '抓側溝--------------------------------------
    m = 0
    Ang = AzToAcadAngle(Pol(LwObj1.Coordinates(1), LwObj1.Coordinates(0), Xy3(1), Xy3(0)))
    '抓起始側側溝--------------------------------------
    AcadDoc.SelectionSets.Item("DhSelectionSetFilter").Delete
    Set objAcadSelectionSet = AcadDoc.SelectionSets.Add("DhSelectionSetFilter")
    Xy2(0) = Xy3(0)
    Xy2(1) = Xy3(1)
    Xy2(2) = 0
    SelectPoint1 = Xy2
    SelectPoint2 = AcadDoc.Utility.PolarPoint(Xy2, Ang - Pi, SearchDist)
    Call RbtSelectCrossing(objAcadSelectionSet, SelectPoint2, SelectPoint1, 2, 0, "LWPOLYLINE", 8, "側溝測線")
    Set LineObj = AcadDoc.ModelSpace.AddLine(SelectPoint1, SelectPoint2)
    LineObj.Color = acBlue
    LineObj.Update
    
    DhCount = 0
    For Each objAcadEntity In objAcadSelectionSet
        If objAcadEntity.Layer = "側溝測線" Then
            Set LwObj2 = objAcadEntity
            varReturn4 = LineObj.IntersectWith(LwObj2, acExtendNone)
            If UBound(varReturn4) >= 2 Then
                DhCount = DhCount + 1
                m = m + 4
                StartP(0) = LwObj2.Coordinates(0)
                StartP(1) = LwObj2.Coordinates(1)
                EndP(0) = LwObj2.Coordinates(UBound(LwObj2.Coordinates) - 1)
                EndP(1) = LwObj2.Coordinates(UBound(LwObj2.Coordinates))
                LwObj2.Color = acRed
                LwObj2.Update
                r = 2
                Do While M_Sheet(7).Cells(r, 1).Value <> 0
                    If Abs(StartP(0) - M_Sheet(7).Cells(r, 3).Value) < 3 And Abs(StartP(1) - M_Sheet(7).Cells(r, 4).Value) < 3 And Abs(EndP(0) - M_Sheet(7).Cells(r, 6).Value) < 3 And Abs(EndP(1) - M_Sheet(7).Cells(r, 7).Value) < 3 Then
                        TransectsDist = Hdist(SelectPoint2(0), SelectPoint2(1), varReturn4(0), varReturn4(1))
                        AvgEL = Round((M_Sheet(7).Cells(r, 5).Value / 2 + M_Sheet(7).Cells(r, 8).Value / 2), 2)
                        AvgDepth = Round((M_Sheet(7).Cells(r, 9).Value / 2 + M_Sheet(7).Cells(r, 10).Value / 2) / 100, 2)
                        AvgWidth = Round((M_Sheet(7).Cells(r, 11).Value / 2 + M_Sheet(7).Cells(r, 12).Value / 2) / 100, 2)
                        M_Sheet(9).Cells(q, 4).Value = Round(TransectsDist, 2)
                        M_Sheet(9).Cells(q, 5).Value = Round(AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2 + 0.01, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2 - 0.01, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL
                        q = q + 1
                    ElseIf Abs(StartP(0) - M_Sheet(7).Cells(r, 6).Value) < 3 And Abs(StartP(1) - M_Sheet(7).Cells(r, 7).Value) < 3 And Abs(EndP(0) - M_Sheet(7).Cells(r, 3).Value) < 3 And Abs(EndP(1) - M_Sheet(7).Cells(r, 4).Value) < 3 Then
                        TransectsDist = Hdist(SelectPoint2(0), SelectPoint2(1), varReturn4(0), varReturn4(1))
                        AvgEL = Round((M_Sheet(7).Cells(r, 5).Value / 2 + M_Sheet(7).Cells(r, 8).Value / 2), 2)
                        AvgDepth = Round((M_Sheet(7).Cells(r, 9).Value / 2 + M_Sheet(7).Cells(r, 10).Value / 2) / 100, 2)
                        AvgWidth = Round((M_Sheet(7).Cells(r, 11).Value / 2 + M_Sheet(7).Cells(r, 12).Value / 2) / 100, 2)
                        M_Sheet(9).Cells(q, 4).Value = Round(TransectsDist, 2)
                        M_Sheet(9).Cells(q, 5).Value = Round(AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2 + 0.01, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2 - 0.01, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL
                        q = q + 1
                    End If
                    r = r + 1
                Loop
                LwObj2.Color = acByLayer
                LwObj2.Update
            End If
        End If
    Next
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
                    For Dh0 = 2 To Dh
                        If Abs(M_Sheet(9).Cells(q, 1).Value - M_Sheet(9).Cells(Dh0, 4).Value) < Abs(M_Sheet(9).Cells(Dh0, 5).Value) Then
                            If M_Sheet(9).Cells(q, 1).Value - M_Sheet(9).Cells(Dh0, 4).Value < 0 Then
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
    Ang = AzToAcadAngle(Pol(LwObj1.Coordinates(UBound(LwObj1.Coordinates)), LwObj1.Coordinates(UBound(LwObj1.Coordinates) - 1), Xy3(1), Xy3(0)))
    '抓終點側側溝--------------------------------------
    AcadDoc.SelectionSets.Item("DhSelectionSetFilter").Delete
    Set objAcadSelectionSet = AcadDoc.SelectionSets.Add("DhSelectionSetFilter")
    Xy2(0) = Xy3(0)
    Xy2(1) = Xy3(1)
    Xy2(2) = 0
    SelectPoint1 = Xy2
    SelectPoint2 = AcadDoc.Utility.PolarPoint(Xy2, Ang - Pi, SearchDist)
    Call RbtSelectCrossing(objAcadSelectionSet, SelectPoint2, SelectPoint1, 2, 0, "LWPOLYLINE", 8, "側溝測線")
    Set LineObj = AcadDoc.ModelSpace.AddLine(SelectPoint1, SelectPoint2)
    LineObj.Color = acBlue
    LineObj.Update
    For Each objAcadEntity In objAcadSelectionSet
        If objAcadEntity.Layer = "側溝測線" Then
            Set LwObj2 = objAcadEntity
            varReturn4 = LineObj.IntersectWith(LwObj2, acExtendNone)
            If UBound(varReturn4) >= 2 Then
                DhCount = DhCount + 1
                m = m + 4
                StartP(0) = LwObj2.Coordinates(0)
                StartP(1) = LwObj2.Coordinates(1)
                EndP(0) = LwObj2.Coordinates(UBound(LwObj2.Coordinates) - 1)
                EndP(1) = LwObj2.Coordinates(UBound(LwObj2.Coordinates))
                LwObj2.Color = acRed
                LwObj2.Update
                r = 2
                Do While M_Sheet(7).Cells(r, 1).Value <> 0
                    If Abs(StartP(0) - M_Sheet(7).Cells(r, 3).Value) < 3 And Abs(StartP(1) - M_Sheet(7).Cells(r, 4).Value) < 3 And Abs(EndP(0) - M_Sheet(7).Cells(r, 6).Value) < 3 And Abs(EndP(1) - M_Sheet(7).Cells(r, 7).Value) < 3 Then
                        TransectsDist = Hdist(SelectPoint3(0), SelectPoint3(1), varReturn4(0), varReturn4(1))
                        AvgEL = Round((M_Sheet(7).Cells(r, 5).Value / 2 + M_Sheet(7).Cells(r, 8).Value / 2), 2)
                        AvgDepth = Round((M_Sheet(7).Cells(r, 9).Value / 2 + M_Sheet(7).Cells(r, 10).Value / 2) / 100, 2)
                        AvgWidth = Round((M_Sheet(7).Cells(r, 11).Value / 2 + M_Sheet(7).Cells(r, 12).Value / 2) / 100, 2)
                        M_Sheet(9).Cells(q, 4).Value = Round(TransectsDist, 2)
                        M_Sheet(9).Cells(q, 5).Value = Round(AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2 + 0.01, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2 - 0.01, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL
                        q = q + 1
                    ElseIf Abs(StartP(0) - M_Sheet(7).Cells(r, 6).Value) < 3 And Abs(StartP(1) - M_Sheet(7).Cells(r, 7).Value) < 3 And Abs(EndP(0) - M_Sheet(7).Cells(r, 3).Value) < 3 And Abs(EndP(1) - M_Sheet(7).Cells(r, 4).Value) < 3 Then
                        TransectsDist = Hdist(SelectPoint3(0), SelectPoint3(1), varReturn4(0), varReturn4(1))
                        AvgEL = Round((M_Sheet(7).Cells(r, 5).Value / 2 + M_Sheet(7).Cells(r, 8).Value / 2), 2)
                        AvgDepth = Round((M_Sheet(7).Cells(r, 9).Value / 2 + M_Sheet(7).Cells(r, 10).Value / 2) / 100, 2)
                        AvgWidth = Round((M_Sheet(7).Cells(r, 11).Value / 2 + M_Sheet(7).Cells(r, 12).Value / 2) / 100, 2)
                        M_Sheet(9).Cells(q, 4).Value = Round(TransectsDist, 2)
                        M_Sheet(9).Cells(q, 5).Value = Round(AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist - AvgWidth / 2 + 0.01, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2 - 0.01, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL - AvgDepth
                        q = q + 1
                        M_Sheet(9).Cells(q, 1).Value = Round(TransectsDist + AvgWidth / 2, 2)
                        M_Sheet(9).Cells(q, 2).Value = AvgEL
                        q = q + 1
                    End If
                    r = r + 1
                Loop
                LwObj2.Color = acByLayer
                LwObj2.Update
            End If
        End If
    Next
    LineObj.Delete
    
 
For Dh0 = Dh + 1 To Dh1
    For Dh00 = Dh2 To q
        If Abs(M_Sheet(9).Cells(Dh0, 1).Value - M_Sheet(9).Cells(Dh00, 4).Value) < Abs(M_Sheet(9).Cells(Dh00, 5).Value) Then
            If M_Sheet(9).Cells(Dh0, 1).Value - M_Sheet(9).Cells(Dh00, 4).Value < 0 Then
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
        .SetRange Range("A2:B" & q - 1)
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
    Do While M_Sheet(9).Cells(u, 1) <> ""
        M_Sheet(4).Cells(w, ww).Value = M_Sheet(9).Cells(u, 2).Value
        ww = ww + 1
        M_Sheet(4).Cells(w, ww).Value = M_Sheet(9).Cells(u, 1).Value
        ww = ww + 1
        u = u + 1
        If ww = 12 And M_Sheet(9).Cells(u, 1).Value <> "" Then
            w = w + 1
            ww = 1
            M_Sheet(4).Cells(w, ww).Value = "GR"
            ww = ww + 1
            M_Sheet(4).Cells(w, ww).Value = M_Sheet(9).Cells(u, 2).Value
            ww = ww + 1
            M_Sheet(4).Cells(w, ww).Value = M_Sheet(9).Cells(u, 1).Value
            ww = ww + 1
            u = u + 1
        End If
    Loop
    w = w + 1
    M_Sheet(4).Cells(w, 1).Value = ";"
    w = w + 1
    M_Sheet(4).Cells(w, 1).Select
    
    LwObj00.Delete
    LwObj01.Delete
    LwObj02.Delete
    LwObj03.Delete
    M_Sheet(9).Activate
    Columns("A:E").Select
    Selection.Delete Shift:=xlToLeft
    M_Sheet(4).Activate
Z:
    M_Sheet(2).Cells(i, 11) = DhCount
    i = i + 1

Loop





End Sub
