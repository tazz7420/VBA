Attribute VB_Name = "G_FixInputOutput"
Sub FixConduits()
'Dim M_Sheet(1 To 9) As Excel.Worksheet
Dim MaxDepth As Double
Dim Depth As Double
Dim Rim As Double
Dim InEL As Double
Dim OutEL As Double

'Set M_Sheet(1) = ThisWorkbook.Worksheets("JUNCTIONS")
'Set M_Sheet(2) = ThisWorkbook.Worksheets("CONDUITS")


i = 1
Do While M_Sheet(2).Cells(i, 1).Value <> ""
    j = 1
    Do While M_Sheet(1).Cells(j, 1).Value <> ""
        If M_Sheet(2).Cells(i, 2).Value = M_Sheet(1).Cells(j, 1).Value Then
            Rim = M_Sheet(1).Cells(j, 2).Value + M_Sheet(1).Cells(j, 3).Value
            M_Sheet(2).Cells(i, 6).Value = Rim - M_Sheet(2).Cells(i, 12).Value
        End If
        j = j + 1
    Loop
    j = 1
    Do While M_Sheet(1).Cells(j, 1).Value <> ""
        If M_Sheet(2).Cells(i, 3).Value = M_Sheet(1).Cells(j, 1).Value Then
            Rim = M_Sheet(1).Cells(j, 2).Value + M_Sheet(1).Cells(j, 3).Value
            M_Sheet(2).Cells(i, 7).Value = Rim - M_Sheet(2).Cells(i, 12).Value
        End If
        j = j + 1
    Loop
    i = i + 1
    M_Sheet(2).Cells(i, 1).Select
Loop



End Sub
