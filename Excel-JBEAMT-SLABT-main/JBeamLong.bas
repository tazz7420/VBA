Attribute VB_Name = "JBeamLong"
Sub ComputeJBeam()
Dim objAcadEntity As AcadEntity
Dim objAcadEntity2 As AcadEntity
Dim objAcadText As AcadText
Dim objAcadText2 As AcadText
Dim objAcadLine As AcadLine
Dim objAcadCircle As AcadCircle
Dim objAcadSelectionSet As AcadSelectionSet
Dim varPoint(1) As Variant
Dim varCenter As Variant
Dim varIntersect As Variant
Dim AlreadyExistJBEAM As Integer
Dim TextRo As Integer
Dim corP1(0 To 2) As Double
Dim JBeamNum As String
Dim LastX, LastY, NewX, NewY As Double

On Error Resume Next
'開啟CAD檔案
'ACADfilename = InputBox("AutoCAD檔案名稱")
'Set AcadDoc = AcadApp.Documents.Open(ThisWorkbook.Path & "\" & AcadFileName)
'AcadDoc.Application.Visible = True
'AcadDoc.WindowState = acMax

ThisWorkbook.Worksheets("JBeam").Cells(1, 1).Value = "檔案名稱:" & AcadFileName
ThisWorkbook.Worksheets("JBeam").Cells(2, 1).Value = Date

AcadDoc.SelectionSets.Item("ComputeJBeam").Delete
Err.Clear

Set objAcadSelectionSet = AcadDoc.SelectionSets.Add("ComputeJBeam")

waitTime = TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + 3)
Application.Wait waitTime
MsgBox ("請於圖面上框選計算範圍")
AcadDoc.Utility.Prompt ("請於圖面上框選計算範圍")
objAcadSelectionSet.SelectOnScreen
LastX = 0
LastY = 0

For Each objAcadEntity In objAcadSelectionSet
    If objAcadEntity.Layer = "JBEAMT" Then
        Set objAcadText = objAcadEntity
        NewX = objAcadText.InsertionPoint(0)
        NewY = objAcadText.InsertionPoint(1)
        If NewX = LastX And NewY = LastY Then
        Else
            LastX = NewX
            LastY = NewY
            ThisWorkbook.Worksheets("JBeam").Cells(k, 1).Value = objAcadText.TextString
            varCenter = objAcadText.InsertionPoint
            Set objAcadCircle = AcadDoc.ModelSpace.AddCircle(varCenter, 15)
            objAcadCircle.Update
            AlreadyExistJBEAM = 0
            For Each objAcadEntity2 In objAcadSelectionSet
                If objAcadEntity2.Layer = "JBEAM" Then
                    Set objAcadLine = objAcadEntity2
                    varIntersect = Empty
                    varIntersect = objAcadCircle.IntersectWith(objAcadLine, acExtendNone)
                    If IsEmpty(varIntersect) Then
                    Else
                        If UBound(varIntersect) < 2 Then
                        Else
                            If AlreadyExistJBEAM = 1 Then
                                TextRo = Round(objAcadText.Rotation * 180 / 3.1415926, 0)
                                If TextRo = 90 And Round(objAcadLine.StartPoint(0), 0) - Round(objAcadLine.EndPoint(0), 0) = 0 Then
                                    ThisWorkbook.Worksheets("JBeam").Cells(k, 2).Value = Round(objAcadLine.Length, 2)
                                ElseIf TextRo = 0 And Round(objAcadLine.StartPoint(1), 0) - Round(objAcadLine.EndPoint(1), 0) = 0 Then
                                    ThisWorkbook.Worksheets("JBeam").Cells(k, 2).Value = Round(objAcadLine.Length, 2)
                                End If
                            Else
                                ThisWorkbook.Worksheets("JBeam").Cells(k, 2).Value = Round(objAcadLine.Length, 2)
                                AlreadyExistJBEAM = 1
                            End If
                            ThisWorkbook.Worksheets("JBeam").Cells(k, 5).Value = "#" & k - 3
                            JBeamNum = "#" & k - 3
                            corP1(0) = objAcadLine.StartPoint(0) / 2 + objAcadLine.EndPoint(0) / 2
                            corP1(1) = objAcadLine.StartPoint(1) / 2 + objAcadLine.EndPoint(1) / 2
                            corP1(2) = 0
                            Set objAcadText2 = AcadDoc.ModelSpace.AddText(JBeamNum, corP1, 20)
                        End If
                    End If
                End If
            Next
            ThisWorkbook.Worksheets("JBeam").Cells(k, 2).Select
            k = k + 1
            objAcadCircle.Delete
            objAcadCircle.Update
        End If
    Else
    End If
Next

End Sub
