Attribute VB_Name = "F_JunctionMaxDepth"
Sub JunctionsMaxDepth()
'Dim M_Sheet(1 To 9) As Excel.Worksheet
Dim MaxDepth As Double
Dim Depth As Double
Dim Rim As Double

'Set M_Sheet(1) = ThisWorkbook.Worksheets("JUNCTIONS")
'Set M_Sheet(2) = ThisWorkbook.Worksheets("CONDUITS")

i = 1
Do While M_Sheet(1).Cells(i, 1).Value <> ""
    MaxDepth = 0
    Rim = M_Sheet(1).Cells(i, 2).Value + M_Sheet(1).Cells(i, 3).Value
    j = 1
    Do While M_Sheet(2).Cells(j, 2).Value <> ""
        If M_Sheet(1).Cells(i, 1).Value = M_Sheet(2).Cells(j, 2).Value Then
            Depth = M_Sheet(2).Cells(j, 12).Value
            If Depth >= MaxDepth Then
                MaxDepth = Depth
                M_Sheet(1).Cells(i, 2).Value = Rim - MaxDepth
                M_Sheet(1).Cells(i, 3).Value = MaxDepth
            End If
        End If
        j = j + 1
    Loop
    j = 1
    Do While M_Sheet(2).Cells(j, 3).Value <> ""
         If M_Sheet(1).Cells(i, 1).Value = M_Sheet(2).Cells(j, 3).Value Then
            Depth = M_Sheet(2).Cells(j, 12).Value
            If Depth >= MaxDepth Then
                MaxDepth = Depth
                M_Sheet(1).Cells(i, 2).Value = Rim - MaxDepth
                M_Sheet(1).Cells(i, 3).Value = MaxDepth
            End If
        End If
        j = j + 1
    Loop
    M_Sheet(1).Activate
    M_Sheet(1).Cells(i, 1).Select
    i = i + 1
Loop




End Sub
