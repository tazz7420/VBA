Attribute VB_Name = "RiverCrossSectionDrawing"
Sub RiverCrossSectionDrawing()
Dim M_sheet(1 To 9) As Excel.Worksheet
Dim AcadApp As New AcadApplication
Dim AcadDoc As AcadDocument
Dim TextObj As AcadText
Dim LwObj As AcadLWPolyline
Dim LwObj2 As AcadLWPolyline
Dim fileName, textValue As String
Dim Xy2(0 To 2) As Double
Dim Xy22(0 To 2) As Double
Dim startPoint(0 To 2) As Double
Dim insertPoint(0 To 2) As Double
Dim Xy3() As Double
Dim Xy5(0 To 5) As Double
Dim Xy4(0 To 3) As Double
Dim minEL, maxEL, totalDistance, startDis2, moveDis As Double
Dim startRow, endRow, hScale, vScale, startDis As Integer

Set M_sheet(1) = ThisWorkbook.Worksheets("Sheet1")

i = 1
Do While M_sheet(1).Cells(i, 1).Value <> ""
    M_sheet(1).Cells(i, 1).Select
    If M_sheet(1).Cells(i, 1).Value = "END" Then
        endRow = i - 1
        ReDim Xy3((endRow - startRow + 1) * 2 - 1)
        i = i + 1
        minEL = Int(minEL)
        vScale = -Int(((maxEL - minEL) / 18) * (-1))
        hScale = -Int((totalDistance / 33) * (-1))
        Xy2(0) = 10: Xy2(1) = 5: Xy2(2) = 0
        Set TextObj = AcadDoc.ModelSpace.AddText("1:" & hScale & "00", Xy2, 3)
        TextObj.Layer = "13": TextObj.Update
        Xy2(0) = 7: Xy2(1) = 20: Xy2(2) = 0
        Set TextObj = AcadDoc.ModelSpace.AddText("1:" & vScale & "00", Xy2, 3)
        TextObj.Alignment = acAlignmentMiddleRight: TextObj.TextAlignmentPoint = Xy2
        TextObj.Rotation = Pi / 2
        TextObj.Layer = "13": TextObj.Update
        ''''''''比例尺
        Xy2(0) = 255: Xy2(1) = 10: Xy2(2) = 0
        For Index = 0 To 5
            Xy2(0) = Xy2(0) + 10
            Set TextObj = AcadDoc.ModelSpace.AddText(Index * hScale, Xy2, 3)
            TextObj.Alignment = acAlignmentCenter: TextObj.TextAlignmentPoint = Xy2
            TextObj.Layer = "16": TextObj.Update
        Next
        Xy2(0) = 255: Xy2(1) = 25: Xy2(2) = 0
        For Index = 0 To 5
            Xy2(0) = Xy2(0) + 10
            Set TextObj = AcadDoc.ModelSpace.AddText(Index * vScale, Xy2, 3)
            TextObj.Alignment = acAlignmentCenter: TextObj.TextAlignmentPoint = Xy2
            TextObj.Layer = "16": TextObj.Update
        Next
        ''''''''''''''
        Xy2(0) = -15: Xy2(1) = -11.5: Xy2(2) = 0
        For Index = 1 To 16
            Xy2(1) = Xy2(1) + 10
            Set TextObj = AcadDoc.ModelSpace.AddText(minEL + vScale * (Index - 1), Xy2, 3)
            TextObj.Layer = "13Y": TextObj.Update
        Next
        Xy2(0) = -10: Xy2(1) = -6: Xy2(2) = 0
        For Index = 1 To 34
            Xy2(0) = Xy2(0) + 10
            Set TextObj = AcadDoc.ModelSpace.AddText(startDis + hScale * (Index - 1), Xy2, 3)
            TextObj.Alignment = acAlignmentCenter: TextObj.TextAlignmentPoint = Xy2
            TextObj.Layer = "13X": TextObj.Update
        Next
        Xy2(0) = 0: Xy2(1) = -53: Xy2(2) = 0
        Set TextObj = AcadDoc.ModelSpace.AddText(fileName, Xy2, 3)
        TextObj.Layer = "14": TextObj.Update
        items = 0
        insertPoint(0) = -3
        moveDis = moveDis / hScale * 10
        For Index = startRow To endRow
            Xy3(items) = moveDis + Hdist(M_sheet(1).Cells(Index, 2).Value, M_sheet(1).Cells(Index, 1).Value, startPoint(1), startPoint(0)) / hScale * 10
            Xy2(0) = Xy3(items): items = items + 1
            Xy3(items) = (M_sheet(1).Cells(Index, 3).Value - minEL) * 10 / vScale
            Xy2(1) = -10: items = items + 1
            If M_sheet(1).Cells(Index, 4).Value = "左樁坐標" Then
                Xy5(0) = Xy2(0)
                Xy5(1) = Xy3(items - 1)
                Xy5(2) = Xy5(0) + 10
                Xy5(3) = Xy5(1) + 10
                Xy5(4) = Xy5(2) + 20: Xy22(0) = Xy5(4)
                Xy5(5) = Xy5(3): Xy22(1) = Xy5(5): Xy22(2) = 0
                Set LwObj2 = AcadDoc.ModelSpace.AddLightWeightPolyline(Xy5)
                LwObj2.Layer = "14"
                LwObj2.Update
                Set TextObj = AcadDoc.ModelSpace.AddText("左斷" & fileName & "  H=" & Format(Round(M_sheet(1).Cells(Index, 3).Value, 2), ".00"), Xy22, 2)
                TextObj.Alignment = acAlignmentMiddleLeft: TextObj.TextAlignmentPoint = Xy22
                TextObj.Layer = "14": TextObj.Update
            End If
            If M_sheet(1).Cells(Index, 4).Value = "右樁坐標" Then
                Xy5(0) = Xy2(0)
                Xy5(1) = Xy3(items - 1)
                Xy5(2) = Xy5(0) - 10
                Xy5(3) = Xy5(1) + 10
                Xy5(4) = Xy5(2) - 20: Xy22(0) = Xy5(4)
                Xy5(5) = Xy5(3): Xy22(1) = Xy5(5): Xy22(2) = 0
                Set LwObj2 = AcadDoc.ModelSpace.AddLightWeightPolyline(Xy5)
                LwObj2.Layer = "14"
                LwObj2.Update
                Set TextObj = AcadDoc.ModelSpace.AddText("右斷" & fileName & "  H=" & Format(Round(M_sheet(1).Cells(Index, 3).Value, 2), ".00"), Xy22, 2)
                TextObj.Alignment = acAlignmentMiddleRight: TextObj.TextAlignmentPoint = Xy22
                TextObj.Layer = "14": TextObj.Update
            End If
            ''''''''''''插入高程&累距
            If Xy2(0) - insertPoint(0) < 2 Then
                Xy4(0) = Xy2(0): Xy4(1) = -44: Xy4(2) = Xy2(0): Xy4(3) = -42
                Set LwObj = AcadDoc.ModelSpace.AddLightWeightPolyline(Xy4)
                LwObj.Layer = "14": LwObj.Update
                Xy4(0) = Xy2(0): Xy4(1) = -26: Xy4(2) = Xy2(0): Xy4(3) = -24
                Set LwObj = AcadDoc.ModelSpace.AddLightWeightPolyline(Xy4)
                LwObj.Layer = "14": LwObj.Update
            Else
                insertPoint(0) = Xy2(0)
                Set TextObj = AcadDoc.ModelSpace.AddText(Format(Round(M_sheet(1).Cells(Index, 3).Value, 2), ".00"), Xy2, 2)
                TextObj.Alignment = acAlignmentMiddleRight: TextObj.TextAlignmentPoint = Xy2
                TextObj.Rotation = Pi / 2
                TextObj.Layer = "14": TextObj.Update
                Xy4(0) = insertPoint(0): Xy4(1) = -26: Xy4(2) = insertPoint(0): Xy4(3) = -24
                Set LwObj = AcadDoc.ModelSpace.AddLightWeightPolyline(Xy4)
                LwObj.Layer = "14": LwObj.Update
                Xy2(1) = -28
                Set TextObj = AcadDoc.ModelSpace.AddText(Format(Round(Hdist(M_sheet(1).Cells(Index, 2).Value, M_sheet(1).Cells(Index, 1).Value, startPoint(1), startPoint(0)) + startDis2, 2), ".00"), Xy2, 2)
                TextObj.Alignment = acAlignmentMiddleRight: TextObj.TextAlignmentPoint = Xy2
                TextObj.Rotation = Pi / 2
                TextObj.Layer = "14": TextObj.Update
                Xy4(0) = insertPoint(0): Xy4(1) = -44: Xy4(2) = insertPoint(0): Xy4(3) = -42
                Set LwObj = AcadDoc.ModelSpace.AddLightWeightPolyline(Xy4)
                LwObj.Layer = "14": LwObj.Update
            End If
            '''''''''''''''''''''''''
        Next
        Set LwObj = AcadDoc.ModelSpace.AddLightWeightPolyline(Xy3)
        LwObj.Layer = "11"
        LwObj.Update
        Xy22(0) = 145
        Xy22(1) = 170
        Xy22(2) = 0
        Set TextObj = AcadDoc.ModelSpace.AddText("急水溪第  " & fileName & "  號斷面", Xy22, 12)
        TextObj.Alignment = acAlignmentMiddleCenter: TextObj.TextAlignmentPoint = Xy22
        TextObj.Layer = "14": TextObj.Update
        AcadDoc.SaveAs ("C:\Users\tazz4\Desktop\Program\零星工作\河道橫斷面繪製\" & fileName)
        AcadDoc.Close
    ElseIf M_sheet(1).Cells(i, 2).Value = "" Then
        fileName = M_sheet(1).Cells(i, 1).Value
        Set AcadDoc = AcadApp.Documents.Open("C:\Users\tazz4\Desktop\Program\零星工作\河道橫斷面繪製\空白斷面.dwg")
        AcadDoc.Application.Visible = True
        AcadDoc.WindowState = acMax
        i = i + 1
        startDis2 = Round(M_sheet(1).Cells(i, 5).Value, 2)
        If M_sheet(1).Cells(i, 5).Value = 0 Then
            startDis = 0
        Else
            startDis = -Int((M_sheet(1).Cells(i, 5).Value) * (-1)) - 1
        End If
        moveDis = M_sheet(1).Cells(i, 5).Value - startDis
        startRow = i
        startPoint(0) = M_sheet(1).Cells(i, 1).Value
        startPoint(1) = M_sheet(1).Cells(i, 2).Value
        startPoint(2) = M_sheet(1).Cells(i, 3).Value: minEL = startPoint(2): maxEL = minEL
    Else
        If CDbl(M_sheet(1).Cells(i, 3).Value) > maxEL Then
            maxEL = M_sheet(1).Cells(i, 3).Value
        End If
        If CDbl(M_sheet(1).Cells(i, 3).Value) < minEL Then
            minEL = M_sheet(1).Cells(i, 3).Value
        End If
        totalDistance = Hdist(M_sheet(1).Cells(i, 2).Value, M_sheet(1).Cells(i, 1).Value, startPoint(1), startPoint(0))
        i = i + 1
    End If
    
Loop

End Sub
