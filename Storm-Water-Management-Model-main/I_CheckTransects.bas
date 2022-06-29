Attribute VB_Name = "I_CheckTransects"
Sub CheckTransects()
Dim ObjCircle(0 To 99) As AcadCircle
Dim ObjLine(0 To 99) As AcadLine
Dim CircleNum, PointNum, LineNum As Integer
Dim Xy2(0 To 2) As Double
Dim Xy22(0 To 2) As Double
Dim FixOrNot, FixOrNot1 As String
Dim objAcadSelectionSet As AcadSelectionSet
Dim objAcadEntity As AcadEntity
Dim FixI, FixJ As Integer


Set AcadDoc1 = AcadApp1.Documents.Open("C:\Users\WXZ\Desktop\Drawing5.dwg")
AcadDoc1.Application.Visible = True
AcadDoc1.WindowState = acMax

i = 1
j = 1
k = 1
CircleNum = 1
LineNum = 1
On Error Resume Next
Do While M_Sheet(4).Cells(i, j).Value <> "END"
    
    If M_Sheet(4).Cells(i, 1).Value = "NC" Then
        i = i + 1
    ElseIf M_Sheet(4).Cells(i, 1).Value = "X1" Then
        Do While M_Sheet(6).Cells(k, 1).Value <> ""
            If M_Sheet(4).Cells(i, 2).Value = M_Sheet(6).Cells(k, 1).Value Then
                Xy2(0) = CDbl(M_Sheet(6).Cells(k, 2).Value)
                Xy2(1) = CDbl(M_Sheet(6).Cells(k, 3).Value)
                Xy2(2) = 0
                Set ObjCircle(CircleNum) = AcadDoc.ModelSpace.AddCircle(Xy2, 5)
                ObjCircle(CircleNum).Update
                CircleNum = CircleNum + 1
                AcadDoc.SendCommand ("zoom" & vbCr & "c" & vbCr & Xy2(0) & "," & Xy2(1) & vbCr & "30" & vbCr)
                Exit Do
            End If
            k = k + 1
        Loop
        i = i + 1
    ElseIf M_Sheet(4).Cells(i, 1).Value = "GR" Then
        j = j + 1
        If j = 12 Then
            j = 1
            i = i + 1
        ElseIf M_Sheet(4).Cells(i, j) = "" Then
            i = i + 1
            GoTo Z
        Else
            Xy2(1) = CDbl(M_Sheet(4).Cells(i, j).Value)
            j = j + 1
            Xy2(0) = CDbl(M_Sheet(4).Cells(i, j).Value)
            Xy2(2) = 0
            PointNum = PointNum + 1
            Set ObjCircle(CircleNum) = AcadDoc1.ModelSpace.AddCircle(Xy2, 0.1)
            ObjCircle(CircleNum).Update
            CircleNum = CircleNum + 1
            If Xy22(0) <> 0 Or Xy22(1) <> 0 Then
                Set ObjLine(LineNum) = AcadDoc1.ModelSpace.AddLine(Xy22, Xy2)
                ObjLine(LineNum).Update
                LineNum = LineNum + 1
            End If
            Xy22(0) = Xy2(0)
            Xy22(1) = Xy2(1)
            Xy22(2) = 0
            'AcadDoc1.SendCommand ("zoom" & vbCr & "c" & vbCr & Xy2(0) & "," & Xy2(1) & vbCr & "20" & vbCr)
        End If
        
    ElseIf M_Sheet(4).Cells(i, 1).Value = ";" Then
Z:
        FixOrNot = "fix"
        FixOrNot1 = InputBox("Fix??")
        If FixOrNot = FixOrNot1 Then

           Stop
            FixI = InputBox("")
            FixII = FixI
            FixJ = 2
            AcadDoc1.SelectionSets.Item("TestSelectionSetFilter").Delete
            Set objAcadSelectionSet = AcadDoc1.SelectionSets.Add("TestSelectionSetFilter")
            objAcadSelectionSet.SelectOnScreen
            For Each objAcadEntity In objAcadSelectionSet
                Set ObjCircle(99) = objAcadEntity
                If FixJ = 12 Then
                    FixJ = 2
                    FixI = FixI + 1
                End If
                M_Sheet(4).Cells(FixI, FixJ).Value = Round(ObjCircle(99).Center(1), 2)
                FixJ = FixJ + 1
                M_Sheet(4).Cells(FixI, FixJ).Value = Round(ObjCircle(99).Center(0), 2)
                FixJ = FixJ + 1
            Next
            Stop
            i = FixII - 2
        End If
        If i Mod 5 = 0 Then
            ThisWorkbook.Save
        End If
        
        'Stop
        'i = 49422
        k = 1
        i = i + 1
        For q = 1 To CircleNum - 1
            ObjCircle(q).Delete
            'ObjCircle(q).Update
        Next
        For q = 1 To LineNum - 1
            ObjLine(q).Delete
            'ObjLine(q).Update
        Next
        CircleNum = 1
        LineNum = 1
        PointNum = 0
        Xy22(0) = 0
        Xy22(1) = 0
        j = 1
    End If
    M_Sheet(4).Cells(i, j).Select
Loop


End Sub
