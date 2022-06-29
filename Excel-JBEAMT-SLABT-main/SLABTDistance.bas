Attribute VB_Name = "SLABTDistance"
Sub ComputeSLAB()
Dim objAcadSelectionSet As AcadSelectionSet
Dim objAcadEntity As AcadEntity
Dim objAcadEntity1 As AcadEntity
Dim objAcadEntity2 As AcadEntity
Dim objAcadText As AcadText
Dim objAcadText2 As AcadText
Dim objAcadLine As AcadLine
Dim objAcadLineA As AcadLine
Dim objAcadLineB As AcadLine
Dim objAcadLineC As AcadLine
Dim objAcadLineD As AcadLine
Dim objAcadLineAA As AcadLine
Dim objAcadLineBB As AcadLine
Dim objAcadLineCC As AcadLine
Dim objAcadLineDD As AcadLine
Dim objAcadLine2 As AcadLine
Dim varCenter As Variant
Dim varIntersect As Variant
Dim varP1, varP2, varP3, varP4, X1, Y1, X2, Y2, FinalDistY, FinalDistX As Variant
Dim PI As Double
Dim DistAY, DistBX, DistCY, DistDX, AY, BX, CY, DX, corAY, corBX, corCY, corDX As Double
Dim corP1(0 To 2) As Double
Dim SLABNum As String
Dim IfCS As String
Dim ChangePointX, ChangePointY As Integer

PI = 4 * Math.Atn(1#)

On Error Resume Next
'開啟CAD檔案
'ACADfilename = InputBox("AutoCAD檔案名稱")
'Set AcadDoc = AcadApp.Documents.Open(ThisWorkbook.Path & "\" & AcadFileName)
'AcadDoc.Application.Visible = True
'AcadDoc.WindowState = acMax

ThisWorkbook.Worksheets("SLAB").Cells(1, 1).Value = "檔案名稱:" & AcadFileName
ThisWorkbook.Worksheets("SLAB").Cells(2, 1).Value = Date

AcadDoc.SelectionSets.Item("ComputeSLAB").Delete
Err.Clear

Set objAcadSelectionSet = AcadDoc.SelectionSets.Add("ComputeSLAB")

waitTime = TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + 3)
Application.Wait waitTime
MsgBox ("請於圖面上框選計算範圍")
AcadDoc.Utility.Prompt ("請於圖面上框選計算範圍")
objAcadSelectionSet.SelectOnScreen

For Each objAcadEntity In objAcadSelectionSet
    If objAcadEntity.Layer = "SLABT" Then
        DistAY = 500
        DistBX = 500
        DistCY = 500
        DistDX = 500
        ChangePointX = 0
        ChangePointY = 0
        Set objAcadText = objAcadEntity
        'ThisWorkbook.Worksheets("JBeam").Cells(k, 1).Value = objAcadText.TextString
        If Left(objAcadText.TextString, 2) = "CS" Then
                
        Else
            varCenter = objAcadText.InsertionPoint
            varP1 = AcadDoc.Utility.PolarPoint(varCenter, PI / 2, 500)
            varP2 = AcadDoc.Utility.PolarPoint(varCenter, 0, 500)
            varP3 = AcadDoc.Utility.PolarPoint(varCenter, -PI / 2, 500)
            varP4 = AcadDoc.Utility.PolarPoint(varCenter, PI, 500)
            Set objAcadLineA = AcadDoc.ModelSpace.AddLine(varCenter, varP1)
            'objAcadLineA.Update
            Set objAcadLineB = AcadDoc.ModelSpace.AddLine(varCenter, varP2)
            'objAcadLineB.Update
            Set objAcadLineC = AcadDoc.ModelSpace.AddLine(varCenter, varP3)
            'objAcadLineC.Update
            Set objAcadLineD = AcadDoc.ModelSpace.AddLine(varCenter, varP4)
            'objAcadLineD.Update
            For Each objAcadEntity1 In objAcadSelectionSet
    '---------------COL-----------------------
                If objAcadEntity1.Layer = "COL" Then
                    Set objAcadLine2 = objAcadEntity1
                    varIntersect = Empty
                    varIntersect = objAcadLineA.IntersectWith(objAcadLine2, acExtendNone)
                    If IsEmpty(varIntersect) Then
                    Else
                        If UBound(varIntersect) < 2 Then
                        Else
                            ChangePointX = 1
                        End If
                    End If
                    varIntersect = Empty
                    varIntersect = objAcadLineB.IntersectWith(objAcadLine2, acExtendNone)
                    If IsEmpty(varIntersect) Then
                    Else
                        If UBound(varIntersect) < 2 Then
                        Else
                            ChangePointY = 2
                        End If
                    End If
                    varIntersect = Empty
                    varIntersect = objAcadLineC.IntersectWith(objAcadLine2, acExtendNone)
                    If IsEmpty(varIntersect) Then
                    Else
                        If UBound(varIntersect) < 2 Then
                        Else
                            ChangePointX = 1
                        End If
                    End If
                    varIntersect = Empty
                    varIntersect = objAcadLineD.IntersectWith(objAcadLine2, acExtendNone)
                    If IsEmpty(varIntersect) Then
                    Else
                        If UBound(varIntersect) < 2 Then
                        Else
                            ChangePointY = 2
                            End If
                        End If
                    End If
            Next
    '---------------JBEAM--------------------
            varCenter = objAcadText.InsertionPoint
            If ChangePointX = 0 And ChangePointY = 0 Then
            ElseIf ChangePointX = 1 And ChangePointY = 2 Then
                varCenter = AcadDoc.Utility.PolarPoint(varCenter, PI / 2, 40)
                varCenter = AcadDoc.Utility.PolarPoint(varCenter, 0, 40)
                objAcadLineA.Delete
                objAcadLineA.Update
                objAcadLineB.Delete
                objAcadLineB.Update
                objAcadLineC.Delete
                objAcadLineC.Update
                objAcadLineD.Delete
                objAcadLineD.Update
                varP1 = AcadDoc.Utility.PolarPoint(varCenter, PI / 2, 500)
                varP2 = AcadDoc.Utility.PolarPoint(varCenter, 0, 500)
                varP3 = AcadDoc.Utility.PolarPoint(varCenter, -PI / 2, 500)
                varP4 = AcadDoc.Utility.PolarPoint(varCenter, PI, 500)
                Set objAcadLineA = AcadDoc.ModelSpace.AddLine(varCenter, varP1)
                'objAcadLineA.Update
                Set objAcadLineB = AcadDoc.ModelSpace.AddLine(varCenter, varP2)
                'objAcadLineB.Update
                Set objAcadLineC = AcadDoc.ModelSpace.AddLine(varCenter, varP3)
                'objAcadLineC.Update
                Set objAcadLineD = AcadDoc.ModelSpace.AddLine(varCenter, varP4)
                'objAcadLineD.Update
            ElseIf ChangePointX = 1 Then
                varCenter = AcadDoc.Utility.PolarPoint(varCenter, 0, 40)
                objAcadLineA.Delete
                objAcadLineA.Update
                objAcadLineB.Delete
                objAcadLineB.Update
                objAcadLineC.Delete
                objAcadLineC.Update
                objAcadLineD.Delete
                objAcadLineD.Update
                varP1 = AcadDoc.Utility.PolarPoint(varCenter, PI / 2, 500)
                varP2 = AcadDoc.Utility.PolarPoint(varCenter, 0, 500)
                varP3 = AcadDoc.Utility.PolarPoint(varCenter, -PI / 2, 500)
                varP4 = AcadDoc.Utility.PolarPoint(varCenter, PI, 500)
                Set objAcadLineA = AcadDoc.ModelSpace.AddLine(varCenter, varP1)
                'objAcadLineA.Update
                Set objAcadLineB = AcadDoc.ModelSpace.AddLine(varCenter, varP2)
                'objAcadLineB.Update
                Set objAcadLineC = AcadDoc.ModelSpace.AddLine(varCenter, varP3)
                'objAcadLineC.Update
                Set objAcadLineD = AcadDoc.ModelSpace.AddLine(varCenter, varP4)
                'objAcadLineD.Update
            ElseIf ChangePointY = 2 Then
                varCenter = AcadDoc.Utility.PolarPoint(varCenter, PI / 2, 40)
                objAcadLineA.Delete
                objAcadLineA.Update
                objAcadLineB.Delete
                objAcadLineB.Update
                objAcadLineC.Delete
                objAcadLineC.Update
                objAcadLineD.Delete
                objAcadLineD.Update
                varP1 = AcadDoc.Utility.PolarPoint(varCenter, PI / 2, 500)
                varP2 = AcadDoc.Utility.PolarPoint(varCenter, 0, 500)
                varP3 = AcadDoc.Utility.PolarPoint(varCenter, -PI / 2, 500)
                varP4 = AcadDoc.Utility.PolarPoint(varCenter, PI, 500)
                Set objAcadLineA = AcadDoc.ModelSpace.AddLine(varCenter, varP1)
                'objAcadLineA.Update
                Set objAcadLineB = AcadDoc.ModelSpace.AddLine(varCenter, varP2)
                'objAcadLineB.Update
                Set objAcadLineC = AcadDoc.ModelSpace.AddLine(varCenter, varP3)
                'objAcadLineC.Update
                Set objAcadLineD = AcadDoc.ModelSpace.AddLine(varCenter, varP4)
                'objAcadLineD.Update
            End If
            For Each objAcadEntity2 In objAcadSelectionSet
                If objAcadEntity2.Layer = "JBEAM" Then
                    Set objAcadLine2 = objAcadEntity2
                    varIntersect = Empty
                    varIntersect = objAcadLineA.IntersectWith(objAcadLine2, acExtendNone)
                    If IsEmpty(varIntersect) Then
                    Else
                        If UBound(varIntersect) < 2 Then
                        Else
                            AY = objAcadLine2.StartPoint(1)
                            If Abs(AY - varCenter(1)) < DistAY Then
                                DistAY = Abs(AY - varCenter(1))
                                corAY = AY
                            End If
                        End If
                    End If
                    varIntersect = Empty
                    varIntersect = objAcadLineB.IntersectWith(objAcadLine2, acExtendNone)
                    If IsEmpty(varIntersect) Then
                    Else
                        If UBound(varIntersect) < 2 Then
                        Else
                            BX = objAcadLine2.StartPoint(0)
                            If Abs(BX - varCenter(0)) < DistBX Then
                                DistBX = Abs(BX - varCenter(0))
                                corBX = BX
                            End If
                        End If
                    End If
                    varIntersect = Empty
                    varIntersect = objAcadLineC.IntersectWith(objAcadLine2, acExtendNone)
                    If IsEmpty(varIntersect) Then
                    Else
                        If UBound(varIntersect) < 2 Then
                        Else
                            CY = objAcadLine2.StartPoint(1)
                            If Abs(CY - varCenter(1)) < DistCY Then
                                DistCY = Abs(CY - varCenter(1))
                                corCY = CY
                            End If
                        End If
                    End If
                    varIntersect = Empty
                    varIntersect = objAcadLineD.IntersectWith(objAcadLine2, acExtendNone)
                    If IsEmpty(varIntersect) Then
                    Else
                        If UBound(varIntersect) < 2 Then
                        Else
                            DX = objAcadLine2.StartPoint(0)
                            If Abs(DX - varCenter(0)) < DistDX Then
                                DistDX = Abs(DX - varCenter(0))
                                corDX = DX
                            End If
                        End If
                    End If
                End If
    '--------------BEAMA--------
                If objAcadEntity2.Layer = "BEAMA" Then
                    Set objAcadLine2 = objAcadEntity2
                    varIntersect = Empty
                    varIntersect = objAcadLineA.IntersectWith(objAcadLine2, acExtendNone)
                    If IsEmpty(varIntersect) Then
                    Else
                        If UBound(varIntersect) < 2 Then
                        Else
                            AY = objAcadLine2.StartPoint(1)
                            If Abs(AY - varCenter(1)) < DistAY Then
                                DistAY = Abs(AY - varCenter(1))
                                corAY = AY
                            End If
                        End If
                    End If
                    varIntersect = Empty
                    varIntersect = objAcadLineB.IntersectWith(objAcadLine2, acExtendNone)
                    If IsEmpty(varIntersect) Then
                    Else
                        If UBound(varIntersect) < 2 Then
                        Else
                            BX = objAcadLine2.StartPoint(0)
                            If Abs(BX - varCenter(0)) < DistBX Then
                                DistBX = Abs(BX - varCenter(0))
                                corBX = BX
                            End If
                        End If
                    End If
                    varIntersect = Empty
                    varIntersect = objAcadLineC.IntersectWith(objAcadLine2, acExtendNone)
                    If IsEmpty(varIntersect) Then
                    Else
                        If UBound(varIntersect) < 2 Then
                        Else
                            CY = objAcadLine2.StartPoint(1)
                            If Abs(CY - varCenter(1)) < DistCY Then
                                DistCY = Abs(CY - varCenter(1))
                                corCY = CY
                            End If
                        End If
                    End If
                    varIntersect = Empty
                    varIntersect = objAcadLineD.IntersectWith(objAcadLine2, acExtendNone)
                    If IsEmpty(varIntersect) Then
                    Else
                        If UBound(varIntersect) < 2 Then
                        Else
                            DX = objAcadLine2.StartPoint(0)
                            If Abs(DX - varCenter(0)) < DistDX Then
                                DistDX = Abs(DX - varCenter(0))
                                corDX = DX
                            End If
                        End If
                    End If
                End If
    '-------------------------------
            Next
        objAcadLineA.Delete
        objAcadLineA.Update
        objAcadLineB.Delete
        objAcadLineB.Update
        objAcadLineC.Delete
        objAcadLineC.Update
        objAcadLineD.Delete
        objAcadLineD.Update
        corP1(0) = corBX / 2 + corDX / 2
        corP1(1) = corAY / 2 + corCY / 2
        corP1(2) = 0
        FinalDistY = Round(DistAY / 2 + DistCY / 2, 2)
        FinalDistX = Round(DistBX / 2 + DistDX / 2, 2)
        varP1 = AcadDoc.Utility.PolarPoint(corP1, PI / 2, FinalDistY)
        varP2 = AcadDoc.Utility.PolarPoint(corP1, 0, FinalDistX)
        varP3 = AcadDoc.Utility.PolarPoint(corP1, -PI / 2, FinalDistY)
        varP4 = AcadDoc.Utility.PolarPoint(corP1, PI, FinalDistX)
        Set objAcadLineAA = AcadDoc.ModelSpace.AddLine(corP1, varP1)
        objAcadLineAA.Update
        Set objAcadLineBB = AcadDoc.ModelSpace.AddLine(corP1, varP2)
        objAcadLineBB.Update
        Set objAcadLineCC = AcadDoc.ModelSpace.AddLine(corP1, varP3)
        objAcadLineCC.Update
        Set objAcadLineDD = AcadDoc.ModelSpace.AddLine(corP1, varP4)
        objAcadLineDD.Update
        ThisWorkbook.Worksheets("SLAB").Cells(q, 1).Value = objAcadText.TextString
        ThisWorkbook.Worksheets("SLAB").Cells(q, 2).Value = Round(DistAY / 2 + DistCY / 2, 2)
        ThisWorkbook.Worksheets("SLAB").Cells(q, 3).Value = Round(DistBX / 2 + DistDX / 2, 2)
        SLABNum = "@" & q - 3
        ThisWorkbook.Worksheets("SLAB").Cells(q, 5).Value = SLABNum
        Set objAcadText2 = AcadDoc.ModelSpace.AddText(SLABNum, corP1, 20)
        q = q + 1
        End If
    End If
Next

End Sub
