Attribute VB_Name = "E_FindMaxDepth"
Sub FindMaxDepth()
'Dim M_Sheet(1 To 9) As Excel.Worksheet
Dim MaxDepth As Double
Dim Depth As Double



'Set M_Sheet(2) = ThisWorkbook.Worksheets("CONDUITS")
'Set M_Sheet(7) = ThisWorkbook.Worksheets("Sheet7")

i = 1

Do While M_Sheet(2).Cells(i, 1).Value <> ""
    j = 13
    MaxDepth = 0
    Do While M_Sheet(2).Cells(i, j).Value <> ""
        k = 1
        
        Do While M_Sheet(7).Cells(k, 1).Value <> ""
            If M_Sheet(2).Cells(i, j).Value = M_Sheet(7).Cells(k, 2).Value Then
                Depth = M_Sheet(7).Cells(k, 9).Value / 2 + M_Sheet(7).Cells(k, 10).Value / 2
                If Depth / 100 > MaxDepth Then
                    MaxDepth = Depth / 100
                    M_Sheet(2).Cells(i, 12).Value = MaxDepth
                End If
            End If
            k = k + 1
        Loop
        j = j + 1
    Loop
    M_Sheet(2).Activate
    M_Sheet(2).Cells(i, 12).Select
    i = i + 1
Loop


End Sub
