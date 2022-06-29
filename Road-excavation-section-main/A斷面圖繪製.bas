Attribute VB_Name = "A斷面圖繪製"
Public Sub 斷面圖繪製()
Dim AcadApp As New AcadApplication
Dim OuterLoop(0 To 0) As AcadEntity
Dim AChatch As AcadHatch
Dim AcadDoc As AcadDocument
Dim i, k As Integer
Dim j As Double
Dim Datprt As Excel.Range
Dim DatCount, intervalCount As Integer
Dim BasePt(0 To 2) As Double
Dim TextPt(0 To 2) As Double
Dim OutXy4(0 To 3) As Double
Dim OutXy6(0 To 5) As Double
Dim tLwObj As AcadLWPolyline
Dim BaseX, intervalX, BaseY, intervalY, Mdistance, PipeB, PipeH, PipeHz, PipeV, PipeBottom, PipeSize As Double
Dim AcadText As AcadText
Dim CirclePipe As AcadCircle
Dim RecPipe(0 To 7) As Double
Dim AcadRecPipe As AcadLWPolyline
Dim Dimpt1(0 To 2) As Double
Dim Dimpt2(0 To 2) As Double
Dim Hpt(0 To 2) As Double
Dim TxtDimPt(0 To 2) As Double
Dim DimLine As AcadLine
Dim ColorNum As Double
Dim PipeType As String
Dim Currentplot As AcadPlot
Dim PipeName As String
Dim PipeNameS As Variant


'計算斷面數量
Set Datprt = ThisWorkbook.Worksheets("工作表1").Cells(1, 1).CurrentRegion
DatCount = Datprt.EntireRow.Count
If DatCount = 1 Then
    MsgBox ("Null")
    GoTo Z
End If


For k = 2 To DatCount
    If ThisWorkbook.Worksheets("工作表1").Cells(k, 1).Value = 0 Then
    GoTo Z
    ElseIf ThisWorkbook.Worksheets("工作表1").Cells(k, 1).Value <> ThisWorkbook.Worksheets("工作表1").Cells(k - 1, 1) Then
    'open AutoCAD
        On Error Resume Next
'        Set Currentplot = AcadDoc.Plot
'        AcadDoc.SetVariable BACKGROUNDPLOT, 0
'
'        AcadDoc.ActiveLayout.ConfigName = "DWG to PDF.pc3" ' Your plot device.
'        AcadDoc.ActiveLayout.CanonicalMediaName = "ISO A3 (420.00 x 297.00 公釐)"
'
'        AcadDoc.ActiveLayout.StandardScale = acScaleToFit
'        AcadDoc.Application.ZoomExtents
'        Currentplot.PlotToDevice
        AcadDoc.SendCommand "-plot" & vbCr & "Y" & vbCr & vbCr & "DWG To PDF.pc3" & vbCr & "ISO A3 (420.00 x 297.00 公釐)" & vbCr & "M" & vbCr & "L" & vbCr & "N" & vbCr & "E" & vbCr & "12.1=1" & vbCr & "C" & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr
'        AcadDoc.SendCommand "-plot" & vbCr & "Y" & vbCr & vbCr & "DWG To PDF.pc3" & vbCr & "w" & vbCr & "M" & vbCr & "L" & vbCr & "N" & vbCr & "w" & vbCr & "0.7615,0.7615" & vbCr & "41.1231,28.9285" & vbCr & "f" & vbCr & "C" & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr
        AcadDoc.Close
        If ThisWorkbook.Worksheets("工作表1").Cells(k, 2).Value <= 20 Then
            Set AcadDoc = AcadApp.Documents.Open(ThisWorkbook.Path & "\PipeBase20.dwg")
        Else
            Set AcadDoc = AcadApp.Documents.Open(ThisWorkbook.Path & "\PipeBase20up.dwg")
        End If
        AcadDoc.Application.Visible = True
        AcadDoc.WindowState = acMax
        AcadDoc.SaveAs ThisWorkbook.Path & "\" & ThisWorkbook.Worksheets("工作表1").Cells(k, 1).Value & ".dwg"
    
    '開始繪製
    '定義起始點
    If ThisWorkbook.Worksheets("工作表1").Cells(k, 2).Value <= 20 Then
        BasePt(0) = 7.4347
        BasePt(1) = 12
        BasePt(2) = 0
    Else
        BasePt(0) = 3.087
        BasePt(1) = 12
        BasePt(2) = 0
    End If
        BaseX = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 2).Value)
        'MsgBox (BaseX)
        '繪製X刻度長
        OutXy4(0) = BasePt(0)
        OutXy4(1) = BasePt(1)
        OutXy4(2) = BasePt(0) + BaseX
        OutXy4(3) = BasePt(1)
        Set tLwObj = AcadDoc.ModelSpace.AddLightWeightPolyline(OutXy4)
        tLwObj.Closed = False
        tLwObj.Update
        '繪製X刻度間距
        intervalX = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 3).Value)
        intervalCount = BaseX / intervalX
        OutXy4(0) = BasePt(0)
        OutXy4(1) = BasePt(1)
        OutXy4(2) = BasePt(0)
        OutXy4(3) = BasePt(1)
        OutXy4(0) = OutXy4(0)
        OutXy4(1) = OutXy4(1) - 0.2
        OutXy4(2) = OutXy4(2)
        OutXy4(3) = OutXy4(3) + 0.2
        Set tLwObj = AcadDoc.ModelSpace.AddLightWeightPolyline(OutXy4)
        tLwObj.Closed = False
        tLwObj.Update
        TextPt(0) = OutXy4(0)
        TextPt(1) = OutXy4(1) - 0.3
        TextPt(2) = 0
        Set AcadText = AcadDoc.ModelSpace.AddText(0, TextPt, 0.2)
        AcadText.StyleName = "標楷體"
        AcadText.Alignment = acAlignmentBottomCenter
        AcadText.TextAlignmentPoint = TextPt
        OutXy4(1) = OutXy4(1) + 0.2
        OutXy4(3) = OutXy4(3) - 0.2
        For i = 1 To intervalCount
        j = i
            If j Mod 10# = 0 Then
                OutXy4(0) = OutXy4(0) + intervalX
                OutXy4(1) = OutXy4(1) - 0.2
                OutXy4(2) = OutXy4(2) + intervalX
                OutXy4(3) = OutXy4(3) + 0.2
                Set tLwObj = AcadDoc.ModelSpace.AddLightWeightPolyline(OutXy4)
                tLwObj.Closed = False
                tLwObj.Update
                TextPt(0) = OutXy4(0)
                TextPt(1) = OutXy4(1) - 0.3
                TextPt(2) = 0
                Set AcadText = AcadDoc.ModelSpace.AddText(intervalX * i, TextPt, 0.2)
                AcadText.StyleName = "標楷體"
                AcadText.Alignment = acAlignmentBottomCenter
                AcadText.TextAlignmentPoint = TextPt
                OutXy4(1) = OutXy4(1) + 0.2
                OutXy4(3) = OutXy4(3) - 0.2
            Else
                OutXy4(0) = OutXy4(0) + intervalX
                OutXy4(1) = OutXy4(1) - 0.1
                OutXy4(2) = OutXy4(2) + intervalX
                OutXy4(3) = OutXy4(3) + 0.1
                Set tLwObj = AcadDoc.ModelSpace.AddLightWeightPolyline(OutXy4)
                tLwObj.Closed = False
                tLwObj.Update
                OutXy4(1) = OutXy4(1) + 0.1
                OutXy4(3) = OutXy4(3) - 0.1
            End If
        Next
        
        '定義起始點
    If ThisWorkbook.Worksheets("工作表1").Cells(k, 2).Value <= 20 Then
        BasePt(0) = 7.4347
        BasePt(1) = 12
        BasePt(2) = 0
    Else
        BasePt(0) = 3.087
        BasePt(1) = 12
        BasePt(2) = 0
    End If
        BaseY = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 4).Value)
        'MsgBox (BaseX)
        '繪製Y刻度長
        OutXy4(0) = BasePt(0)
        OutXy4(1) = BasePt(1)
        OutXy4(2) = BasePt(0)
        OutXy4(3) = BasePt(1) + Abs(BaseY)
        Set tLwObj = AcadDoc.ModelSpace.AddLightWeightPolyline(OutXy4)
        tLwObj.Closed = False
        tLwObj.Update
        '繪製Y刻度間距
        intervalY = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 5).Value)
        intervalCount = Abs(BaseY) / intervalY
        OutXy4(0) = BasePt(0)
        OutXy4(1) = BasePt(1) + Abs(BaseY)
        OutXy4(2) = BasePt(0)
        OutXy4(3) = BasePt(1) + Abs(BaseY)
        OutXy4(0) = OutXy4(0) - 0.2
        OutXy4(1) = OutXy4(1)
        OutXy4(2) = OutXy4(2) + 0.2
        OutXy4(3) = OutXy4(3)
        Set tLwObj = AcadDoc.ModelSpace.AddLightWeightPolyline(OutXy4)
        tLwObj.Closed = False
        tLwObj.Update
        TextPt(0) = OutXy4(0) - 0.3
        TextPt(1) = OutXy4(1)
        TextPt(2) = 0
        If BaseY > 0 Then
            Set AcadText = AcadDoc.ModelSpace.AddText(0, TextPt, 0.2)
        Else
            Set AcadText = AcadDoc.ModelSpace.AddText(Abs(BaseY), TextPt, 0.2)
        End If
        AcadText.StyleName = "標楷體"
        AcadText.Alignment = acAlignmentLeft
        OutXy4(0) = OutXy4(0) + 0.2
        OutXy4(2) = OutXy4(2) - 0.2
        For i = 1 To intervalCount
        j = i
            If j Mod 10# = 0 Then
                OutXy4(0) = OutXy4(0) - 0.2
                OutXy4(1) = OutXy4(1) - intervalY
                OutXy4(2) = OutXy4(2) + 0.2
                OutXy4(3) = OutXy4(3) - intervalY
                Set tLwObj = AcadDoc.ModelSpace.AddLightWeightPolyline(OutXy4)
                tLwObj.Closed = False
                tLwObj.Update
                TextPt(0) = OutXy4(0) - 0.3
                TextPt(1) = OutXy4(1)
                TextPt(2) = 0
                If BaseY < 0 Then
                    Set AcadText = AcadDoc.ModelSpace.AddText(intervalCount / 10 - i / 10, TextPt, 0.2)
                Else
                    Set AcadText = AcadDoc.ModelSpace.AddText(intervalY * i, TextPt, 0.2)
                End If
                AcadText.StyleName = "標楷體"
                AcadText.Alignment = acAlignmentLeft
                OutXy4(0) = OutXy4(0) + 0.2
                OutXy4(2) = OutXy4(2) - 0.2
            Else
                OutXy4(0) = OutXy4(0) - 0.1
                OutXy4(1) = OutXy4(1) - intervalY
                OutXy4(2) = OutXy4(2) + 0.1
                OutXy4(3) = OutXy4(3) - intervalY
                Set tLwObj = AcadDoc.ModelSpace.AddLightWeightPolyline(OutXy4)
                tLwObj.Closed = False
                tLwObj.Update
                OutXy4(0) = OutXy4(0) + 0.1
                OutXy4(2) = OutXy4(2) - 0.1
            End If
        Next
        '測量長度繪製
        '定義起始點
    If ThisWorkbook.Worksheets("工作表1").Cells(k, 2).Value <= 20 Then
        BasePt(0) = 7.4347
        BasePt(1) = 12
        BasePt(2) = 0
    Else
        BasePt(0) = 3.087
        BasePt(1) = 12
        BasePt(2) = 0
    End If
        Mdistance = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 6).Value)
        '繪製測量長度
        OutXy6(0) = BasePt(0)
        OutXy6(1) = BasePt(1) + Abs(BaseY)
        OutXy6(2) = BasePt(0) + Mdistance
        OutXy6(3) = BasePt(1) + Abs(BaseY)
        OutXy6(4) = BasePt(0) + Mdistance
        OutXy6(5) = BasePt(1)
        Set tLwObj = AcadDoc.ModelSpace.AddLightWeightPolyline(OutXy6)
        tLwObj.Closed = False
        tLwObj.Update
        '標註的長度繪製
        Dimpt1(0) = BasePt(0)
        Dimpt1(1) = BasePt(1) + Abs(BaseY) + 0.1
        Dimpt2(0) = BasePt(0)
        Dimpt2(1) = BasePt(1) + Abs(BaseY) + 1.2
        Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
        DimLine.Update
        Dimpt1(0) = BasePt(0) + Mdistance
        Dimpt1(1) = BasePt(1) + Abs(BaseY) + 0.1
        Dimpt2(0) = BasePt(0) + Mdistance
        Dimpt2(1) = BasePt(1) + Abs(BaseY) + 1.2
        Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
        DimLine.Update
        Dimpt1(0) = BasePt(0)
        Dimpt1(1) = BasePt(1) + Abs(BaseY) + 0.8
        Dimpt2(0) = BasePt(0) + Mdistance
        Dimpt2(1) = BasePt(1) + Abs(BaseY) + 0.8
        Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
        DimLine.Update
        Dimpt1(0) = BasePt(0)
        Dimpt1(1) = BasePt(1) + Abs(BaseY) + 0.8
        Dimpt2(0) = BasePt(0) + 0.1
        Dimpt2(1) = BasePt(1) + Abs(BaseY) + 1#
        Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
        DimLine.Update
        Dimpt1(0) = BasePt(0)
        Dimpt1(1) = BasePt(1) + Abs(BaseY) + 0.8
        Dimpt2(0) = BasePt(0) + 0.1
        Dimpt2(1) = BasePt(1) + Abs(BaseY) + 0.6
        Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
        DimLine.Update
        Dimpt1(0) = BasePt(0) + Mdistance
        Dimpt1(1) = BasePt(1) + Abs(BaseY) + 0.8
        Dimpt2(0) = BasePt(0) + Mdistance - 0.1
        Dimpt2(1) = BasePt(1) + Abs(BaseY) + 0.6
        Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
        DimLine.Update
        Dimpt1(0) = BasePt(0) + Mdistance
        Dimpt1(1) = BasePt(1) + Abs(BaseY) + 0.8
        Dimpt2(0) = BasePt(0) + Mdistance - 0.1
        Dimpt2(1) = BasePt(1) + Abs(BaseY) + 1#
        Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
        DimLine.Update
        PipeNameS = Split(ThisWorkbook.Worksheets("工作表1").Cells(k, 1).Value, "-")
        
        '標註的文字繪製
        '定義文字點
    If ThisWorkbook.Worksheets("工作表1").Cells(k, 2).Value <= 20 Then
        BasePt(0) = 7.4347
        BasePt(1) = 12
        BasePt(2) = 0
    Else
        BasePt(0) = 3.087
        BasePt(1) = 12
        BasePt(2) = 0
    End If
        TextPt(0) = BasePt(0) + BaseX / 2
        TextPt(1) = BasePt(1) + Abs(BaseY) + 3
        PipeName = ThisWorkbook.Worksheets("工作表1").Cells(k, 1).Value
        Set AcadText = AcadDoc.ModelSpace.AddText("斷面" & PipeName, TextPt, 0.8)
        AcadText.StyleName = "標楷體"
        AcadText.Alignment = acAlignmentBottomCenter
        AcadText.TextAlignmentPoint = TextPt
        AcadText.Update
    If ThisWorkbook.Worksheets("工作表1").Cells(k, 2).Value <= 20 Then
        BasePt(0) = 7.4347
        BasePt(1) = 12
        BasePt(2) = 0
    Else
        BasePt(0) = 3.087
        BasePt(1) = 12
        BasePt(2) = 0
    End If
        TextPt(0) = BasePt(0) + BaseX / 2
        TextPt(1) = BasePt(1) + Abs(BaseY) + 2.6
        PipeName = ThisWorkbook.Worksheets("工作表1").Cells(k, 6).Value
        Set AcadText = AcadDoc.ModelSpace.AddText("測量距離" & PipeName, TextPt, 0.4)
        AcadText.StyleName = "標楷體"
        AcadText.Alignment = acAlignmentBottomCenter
        AcadText.TextAlignmentPoint = TextPt
        AcadText.Update
    If ThisWorkbook.Worksheets("工作表1").Cells(k, 2).Value <= 20 Then
        BasePt(0) = 7.4347
        BasePt(1) = 12
        BasePt(2) = 0
    Else
        BasePt(0) = 3.087
        BasePt(1) = 12
        BasePt(2) = 0
    End If
        TextPt(0) = BasePt(0)
        TextPt(1) = BasePt(1) + Abs(BaseY) + 1.5
        PipeName = ThisWorkbook.Worksheets("工作表1").Cells(k, 6).Value
        Set AcadText = AcadDoc.ModelSpace.AddText(PipeNameS(0), TextPt, 0.4)
        AcadText.StyleName = "標楷體"
        AcadText.Alignment = acAlignmentBottomCenter
        AcadText.TextAlignmentPoint = TextPt
        AcadText.Update
    If ThisWorkbook.Worksheets("工作表1").Cells(k, 2).Value <= 20 Then
        BasePt(0) = 7.4347
        BasePt(1) = 12
        BasePt(2) = 0
    Else
        BasePt(0) = 3.087
        BasePt(1) = 12
        BasePt(2) = 0
    End If
        TextPt(0) = BasePt(0) + Mdistance
        TextPt(1) = BasePt(1) + Abs(BaseY) + 1.5
        PipeName = ThisWorkbook.Worksheets("工作表1").Cells(k, 6).Value
        Set AcadText = AcadDoc.ModelSpace.AddText(PipeNameS(1), TextPt, 0.4)
        AcadText.StyleName = "標楷體"
        AcadText.Alignment = acAlignmentBottomCenter
        AcadText.TextAlignmentPoint = TextPt
        AcadText.Update
        
        '繪製圖框標註
        TextPt(0) = 32
        TextPt(1) = 6
        PipeName = ThisWorkbook.Worksheets("工作表1").Cells(k, 1).Value
        Set AcadText = AcadDoc.ModelSpace.AddText(PipeName, TextPt, 0.4)
        AcadText.StyleName = "標楷體"
        AcadText.Alignment = acAlignmentBottomCenter
        AcadText.TextAlignmentPoint = TextPt
        
        
        '繪製管線位置
        '定義起始點
If BaseY > 0 Then
    If ThisWorkbook.Worksheets("工作表1").Cells(k, 2).Value <= 20 Then
        BasePt(0) = 7.4347
        BasePt(1) = 12
        BasePt(2) = 0
    Else
        BasePt(0) = 3.087
        BasePt(1) = 12
        BasePt(2) = 0
    End If
        PipeType = ThisWorkbook.Worksheets("工作表1").Cells(k, 7).Value
        PipeB = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 9).Value)
        PipeH = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 10).Value)
        PipeHz = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 11).Value)
        PipeV = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 12).Value)
        PipeBottom = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 14).Value)
        PipeSize = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 15).Value)
        ColorNum = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 13).Value)
        '圓形管
        If PipeH = 0 Then
        If ThisWorkbook.Worksheets("工作表1").Cells(k, 2).Value <= 20 Then
            BasePt(0) = PipeHz + 7.4347
            BasePt(1) = Abs(BaseY) - PipeV - PipeB / 2000 + 12
            BasePt(2) = 0
        Else
            BasePt(0) = PipeHz + 3.087
            BasePt(1) = Abs(BaseY) - PipeV - PipeB / 2000 + 12
        End If
            Set CirclePipe = AcadDoc.ModelSpace.AddCircle(BasePt, Abs(PipeB) / 1000 / 2)
            CirclePipe.Color = ColorNum
            CirclePipe.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = Abs(BaseY) + 12
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 12
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.LineType = "HIDDEN"
''''''''''''''''''''''''''''''''標住的小短線''''''''''''''''''''''''''''''
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 10.66
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 10.46
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 9.36
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 9.56
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 8.92
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 8.72
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 7.81
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 7.61
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 3.68
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 3.48
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 2.57
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 2.37
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 7.14
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 6.94
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 6.06
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 5.86
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 5.43
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 5.23
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 4.32
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 4.12
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            TextPt(0) = BasePt(0) - PipeB / 2000 - 0.3
            TextPt(1) = BasePt(1)
            Set AcadText = AcadDoc.ModelSpace.AddText(PipeType, TextPt, 0.2)
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentLeft
            AcadText.Color = ColorNum
            AcadText.Update
            '管線名稱
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 10
            Set AcadText = AcadDoc.ModelSpace.AddText(PipeType, TextPt, 0.2)
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
            '管頂深
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 8.3
            If PipeB < 0 Then
                Set AcadText = AcadDoc.ModelSpace.AddText(PipeV + PipeB / 1000, TextPt, 0.2)
            Else
                Set AcadText = AcadDoc.ModelSpace.AddText(PipeV, TextPt, 0.2)
            End If
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
            '管底深
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 6.52
            If PipeB < 0 Then
                Set AcadText = AcadDoc.ModelSpace.AddText(PipeBottom, TextPt, 0.2)
            Else
                Set AcadText = AcadDoc.ModelSpace.AddText(PipeBottom, TextPt, 0.2)
            End If
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
            '管徑
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 4.77
            Set AcadText = AcadDoc.ModelSpace.AddText(PipeB, TextPt, 0.2)
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
            '距離
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 3
            Set AcadText = AcadDoc.ModelSpace.AddText(PipeHz, TextPt, 0.2)
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
        '矩形管
        Else
        If ThisWorkbook.Worksheets("工作表1").Cells(k, 2).Value <= 20 Then
            BasePt(0) = PipeHz + 7.4347
            BasePt(1) = Abs(BaseY) - PipeV - PipeH / 2000 + 12
            BasePt(2) = 0
        Else
            BasePt(0) = PipeHz + 3.087
            BasePt(1) = Abs(BaseY) - PipeV - PipeH / 2000 + 12
        End If
            RecPipe(0) = BasePt(0) - PipeB / 1000 / 2
            RecPipe(1) = BasePt(1) - PipeH / 1000 / 2
            RecPipe(2) = BasePt(0) + PipeB / 1000 / 2
            RecPipe(3) = BasePt(1) - PipeH / 1000 / 2
            RecPipe(4) = BasePt(0) + PipeB / 1000 / 2
            RecPipe(5) = BasePt(1) + PipeH / 1000 / 2
            RecPipe(6) = BasePt(0) - PipeB / 1000 / 2
            RecPipe(7) = BasePt(1) + PipeH / 1000 / 2
            Set AcadRecPipe = AcadDoc.ModelSpace.AddLightWeightPolyline(RecPipe)
            AcadRecPipe.Color = ColorNum
            AcadRecPipe.Closed = True
            AcadRecPipe.Update
            If PipeV = 0 Then
            Set OuterLoop(0) = AcadRecPipe
            Set AChatch = AcadDoc.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True)
            AChatch.AppendOuterLoop OuterLoop
            AChatch.Evaluate
            AChatch.Color = ColorNum
            AChatch.Update
            Else
            End If
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = Abs(BaseY) + 12
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 12
            If PipeV = 0 Then
            Else
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.LineType = "HIDDEN"
            DimLine.Update
            End If
            If PipeV = 0 Then
            TextPt(0) = BasePt(0) - PipeB / 2000 - 0.5
            TextPt(1) = BasePt(1) + 0.2
            Else
            TextPt(0) = BasePt(0) - PipeB / 2000 - 0.3
            TextPt(1) = BasePt(1)
            End If
            If PipeV = 0 Then
            Set AcadText = AcadDoc.ModelSpace.AddText(PipeType & "(" & PipeHz & ")", TextPt, 0.2)
            Else
            Set AcadText = AcadDoc.ModelSpace.AddText(PipeType, TextPt, 0.2)
            End If
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentLeft
            AcadText.Color = ColorNum
            AcadText.Update
            If PipeV = 0 Then
            Else
''''''''''''''''''''''''''''''''標住的小短線''''''''''''''''''''''''''''''
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 10.66
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 10.46
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 9.36
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 9.56
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 8.92
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 8.72
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 7.81
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 7.61
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 3.68
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 3.48
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 2.57
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 2.37
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 7.14
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 6.94
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 6.06
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 5.86
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 5.43
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 5.23
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 4.32
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 4.12
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            End If
            '管線名稱
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 10
            Set AcadText = AcadDoc.ModelSpace.AddText(PipeType, TextPt, 0.2)
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
            '管頂深
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 8.3
            If PipeH < 0 Then
                Set AcadText = AcadDoc.ModelSpace.AddText(PipeV + PipeH / 1000, TextPt, 0.2)
            Else
                Set AcadText = AcadDoc.ModelSpace.AddText(PipeV, TextPt, 0.2)
            End If
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
            '管底深
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 6.52
            If PipeH < 0 Then
                Set AcadText = AcadDoc.ModelSpace.AddText(PipeBottom, TextPt, 0.2)
            Else
                Set AcadText = AcadDoc.ModelSpace.AddText(PipeBottom, TextPt, 0.2)
            End If
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
            '管徑
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 4.77
            Set AcadText = AcadDoc.ModelSpace.AddText(Abs(PipeB) & "x" & Abs(PipeH), TextPt, 0.1)
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
            '距離
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 3
            Set AcadText = AcadDoc.ModelSpace.AddText(PipeHz, TextPt, 0.2)
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
        End If
Else
'        BasePt(0) = 9
'        BasePt(1) = 12
'        BasePt(2) = 0
'        PipeType = ThisWorkbook.Worksheets("工作表1").Cells(k, 7).Value
'        PipeB = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 9).Value)
'        PipeH = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 10).Value)
'        PipeHz = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 11).Value)
'        PipeV = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 12).Value)
'        ColorNum = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 13).Value)
'        '圓形管
'        If PipeH = 0 Then
'            BasePt(0) = PipeHz + 9
'            BasePt(1) = -PipeV + 12
'            BasePt(2) = 0
'            Set CirclePipe = AcadDoc.ModelSpace.AddCircle(BasePt, Abs(PipeB) / 1000 / 2)
'            CirclePipe.Color = ColorNum
'            CirclePipe.Update
'            Dimpt1(0) = BasePt(0)
'            Dimpt1(1) = Abs(BaseY) + 12
'            Dimpt2(0) = BasePt(0)
'            Dimpt2(1) = 12
'            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
'            DimLine.LineType = "HIDDEN"
'            DimLine.Update
'            TextPt(0) = BasePt(0) - PipeB / 2000 - 0.3
'            TextPt(1) = BasePt(1)
'            Set AcadText = AcadDoc.ModelSpace.AddText(PipeType, TextPt, 0.2)
'            AcadText.Alignment = acAlignmentLeft
'            AcadText.Color = ColorNum
'            AcadText.Update
'        '矩形管
'        Else
'            BasePt(0) = PipeHz + 9
'            BasePt(1) = -PipeV + 12
'            BasePt(2) = 0
'            RecPipe(0) = BasePt(0) - PipeB / 1000 / 2
'            RecPipe(1) = BasePt(1) - PipeH / 1000 / 2
'            RecPipe(2) = BasePt(0) + PipeB / 1000 / 2
'            RecPipe(3) = BasePt(1) - PipeH / 1000 / 2
'            RecPipe(4) = BasePt(0) + PipeB / 1000 / 2
'            RecPipe(5) = BasePt(1) + PipeH / 1000 / 2
'            RecPipe(6) = BasePt(0) - PipeB / 1000 / 2
'            RecPipe(7) = BasePt(1) + PipeH / 1000 / 2
'            Set AcadRecPipe = AcadDoc.ModelSpace.AddLightWeightPolyline(RecPipe)
'            AcadRecPipe.Color = ColorNum
'            AcadRecPipe.Closed = True
'            AcadRecPipe.Update
'            Dimpt1(0) = BasePt(0)
'            Dimpt1(1) = Abs(BaseY) + 12
'            Dimpt2(0) = BasePt(0)
'            Dimpt2(1) = 12
'            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
'            DimLine.LineType = "HIDDEN"
'            DimLine.Update
'            TextPt(0) = BasePt(0) - PipeB / 2000 - 0.3
'            TextPt(1) = BasePt(1)
'            Set AcadText = AcadDoc.ModelSpace.AddText(PipeType, TextPt, 0.2)
'            AcadText.Alignment = acAlignmentLeft
'            AcadText.Color = ColorNum
'            AcadText.Update
'        End If
End If
    Else
If BaseY > 0 Then
    If ThisWorkbook.Worksheets("工作表1").Cells(k, 2).Value <= 20 Then
        BasePt(0) = 7.4347
        BasePt(1) = 12
        BasePt(2) = 0
    Else
        BasePt(0) = 3.087
        BasePt(1) = 12
        BasePt(2) = 0
    End If
        PipeType = ThisWorkbook.Worksheets("工作表1").Cells(k, 7).Value
        PipeB = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 9).Value)
        PipeH = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 10).Value)
        PipeHz = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 11).Value)
        PipeV = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 12).Value)
        PipeBottom = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 14).Value)
        PipeSize = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 15).Value)
        ColorNum = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 13).Value)
        '圓形管
        If PipeH = 0 Then
        If ThisWorkbook.Worksheets("工作表1").Cells(k, 2).Value <= 20 Then
            BasePt(0) = PipeHz + 7.4347
            BasePt(1) = Abs(BaseY) - PipeV - PipeB / 2000 + 12
            BasePt(2) = 0
        Else
            BasePt(0) = PipeHz + 3.087
            BasePt(1) = Abs(BaseY) - PipeV - PipeB / 2000 + 12
        End If
            Set CirclePipe = AcadDoc.ModelSpace.AddCircle(BasePt, Abs(PipeB) / 1000 / 2)
            CirclePipe.Color = ColorNum
            CirclePipe.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = Abs(BaseY) + 12
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 12
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.LineType = "HIDDEN"
''''''''''''''''''''''''''''''''標住的小短線''''''''''''''''''''''''''''''
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 10.66
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 10.46
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 9.36
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 9.56
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 8.92
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 8.72
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 7.81
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 7.61
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 3.68
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 3.48
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 2.57
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 2.37
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 7.14
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 6.94
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 6.06
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 5.86
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 5.43
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 5.23
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 4.32
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 4.12
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            TextPt(0) = BasePt(0) - PipeB / 2000 - 0.3
            TextPt(1) = BasePt(1)
            Set AcadText = AcadDoc.ModelSpace.AddText(PipeType, TextPt, 0.2)
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentLeft
            AcadText.Color = ColorNum
            AcadText.Update
            '管線名稱
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 10
            Set AcadText = AcadDoc.ModelSpace.AddText(PipeType, TextPt, 0.2)
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
            '管頂深
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 8.3
            If PipeB < 0 Then
                Set AcadText = AcadDoc.ModelSpace.AddText(PipeV + PipeB / 1000, TextPt, 0.2)
            Else
                Set AcadText = AcadDoc.ModelSpace.AddText(PipeV, TextPt, 0.2)
            End If
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
            '管底深
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 6.52
            If PipeB < 0 Then
                Set AcadText = AcadDoc.ModelSpace.AddText(PipeBottom, TextPt, 0.2)
            Else
                Set AcadText = AcadDoc.ModelSpace.AddText(PipeBottom, TextPt, 0.2)
            End If
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
            '管徑
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 4.77
            Set AcadText = AcadDoc.ModelSpace.AddText(PipeB, TextPt, 0.2)
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
            '距離
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 3
            Set AcadText = AcadDoc.ModelSpace.AddText(PipeHz, TextPt, 0.2)
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
        '矩形管
        Else
        If ThisWorkbook.Worksheets("工作表1").Cells(k, 2).Value <= 20 Then
            BasePt(0) = PipeHz + 7.4347
            BasePt(1) = Abs(BaseY) - PipeV - PipeH / 2000 + 12
            BasePt(2) = 0
        Else
            BasePt(0) = PipeHz + 3.087
            BasePt(1) = Abs(BaseY) - PipeV - PipeH / 2000 + 12
        End If
            RecPipe(0) = BasePt(0) - PipeB / 1000 / 2
            RecPipe(1) = BasePt(1) - PipeH / 1000 / 2
            RecPipe(2) = BasePt(0) + PipeB / 1000 / 2
            RecPipe(3) = BasePt(1) - PipeH / 1000 / 2
            RecPipe(4) = BasePt(0) + PipeB / 1000 / 2
            RecPipe(5) = BasePt(1) + PipeH / 1000 / 2
            RecPipe(6) = BasePt(0) - PipeB / 1000 / 2
            RecPipe(7) = BasePt(1) + PipeH / 1000 / 2
            Set AcadRecPipe = AcadDoc.ModelSpace.AddLightWeightPolyline(RecPipe)
            AcadRecPipe.Color = ColorNum
            AcadRecPipe.Closed = True
            AcadRecPipe.Update
            If PipeV = 0 Then
            Set OuterLoop(0) = AcadRecPipe
            Set AChatch = AcadDoc.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True)
            AChatch.AppendOuterLoop OuterLoop
            AChatch.Evaluate
            AChatch.Color = ColorNum
            AChatch.Update
'            Hpt(0) = (RecPipe(0) + RecPipe(2)) / 2
'            Hpt(1) = (RecPipe(1) + RecPipe(3)) / 2
'            AcadDoc.SendCommand "H" & vbCr & "k" & vbCr & Hpt & vbCr & vbCr
            Else
            End If
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = Abs(BaseY) + 12
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 12
            If PipeV = 0 Then
            Else
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.LineType = "HIDDEN"
            DimLine.Update
            End If
            If PipeV = 0 Then
            TextPt(0) = BasePt(0) - PipeB / 2000 - 0.5
            TextPt(1) = BasePt(1) + 0.2
            Else
            TextPt(0) = BasePt(0) - PipeB / 2000 - 0.3
            TextPt(1) = BasePt(1)
            End If
            If PipeV = 0 Then
            Set AcadText = AcadDoc.ModelSpace.AddText(PipeType & "(" & PipeHz & ")", TextPt, 0.2)
            Else
            Set AcadText = AcadDoc.ModelSpace.AddText(PipeType, TextPt, 0.2)
            End If
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentLeft
            AcadText.Color = ColorNum
            AcadText.Update
            If PipeV = 0 Then
            Else
''''''''''''''''''''''''''''''''標住的小短線''''''''''''''''''''''''''''''
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 10.66
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 10.46
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 9.36
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 9.56
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 8.92
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 8.72
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 7.81
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 7.61
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 3.68
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 3.48
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 2.57
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 2.37
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 7.14
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 6.94
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 6.06
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 5.86
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 5.43
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 5.23
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
            Dimpt1(0) = BasePt(0)
            Dimpt1(1) = 4.32
            Dimpt2(0) = BasePt(0)
            Dimpt2(1) = 4.12
            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
            DimLine.Update
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            End If
            '管線名稱
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 10
            Set AcadText = AcadDoc.ModelSpace.AddText(PipeType, TextPt, 0.2)
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
            '管頂深
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 8.3
            If PipeH < 0 Then
                Set AcadText = AcadDoc.ModelSpace.AddText(PipeV + PipeH / 1000, TextPt, 0.2)
            Else
                Set AcadText = AcadDoc.ModelSpace.AddText(PipeV, TextPt, 0.2)
            End If
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
            '管底深
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 6.52
            If PipeH < 0 Then
                Set AcadText = AcadDoc.ModelSpace.AddText(PipeBottom, TextPt, 0.2)
            Else
                Set AcadText = AcadDoc.ModelSpace.AddText(PipeBottom, TextPt, 0.2)
            End If
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
            '管徑
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 4.77
            Set AcadText = AcadDoc.ModelSpace.AddText(Abs(PipeB) & "x" & Abs(PipeH), TextPt, 0.1)
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
            '距離
            If PipeV = 0 Then
            Else
            TextPt(0) = BasePt(0)
            TextPt(1) = 3
            Set AcadText = AcadDoc.ModelSpace.AddText(PipeHz, TextPt, 0.2)
            AcadText.StyleName = "標楷體"
            AcadText.Alignment = acAlignmentMiddleCenter
            AcadText.TextAlignmentPoint = TextPt
            AcadText.Rotate TextPt, 3.141592653589 / 2
            AcadText.Color = ColorNum
            AcadText.Update
            End If
        End If
Else
'        '繪製管線位置
'        '定義起始點
'        BasePt(0) = 9
'        BasePt(1) = 12
'        BasePt(2) = 0
'        PipeType = ThisWorkbook.Worksheets("工作表1").Cells(k, 7).Value
'        PipeB = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 9).Value)
'        PipeH = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 10).Value)
'        PipeHz = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 11).Value)
'        PipeV = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 12).Value)
'        ColorNum = Val(ThisWorkbook.Worksheets("工作表1").Cells(k, 13).Value)
'        '圓形管
'        If PipeH = 0 Then
'            BasePt(0) = PipeHz + 9
'            BasePt(1) = -PipeV + 12
'            BasePt(2) = 0
'            Set CirclePipe = AcadDoc.ModelSpace.AddCircle(BasePt, PipeB / 1000 / 2)
'            CirclePipe.Color = ColorNum
'            CirclePipe.Update
'            Dimpt1(0) = BasePt(0)
'            Dimpt1(1) = Abs(BaseY) + 12
'            Dimpt2(0) = BasePt(0)
'            Dimpt2(1) = 12
'            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
'            DimLine.LineType = "HIDDEN"
'            DimLine.Update
'            TextPt(0) = BasePt(0) - Abs(PipeB) / 2000 - 0.3
'            TextPt(1) = BasePt(1)
'            Set AcadText = AcadDoc.ModelSpace.AddText(PipeType, TextPt, 0.2)
'            AcadText.Alignment = acAlignmentLeft
'            AcadText.Color = ColorNum
'            AcadText.Update
'        '矩形管
'        Else
'            BasePt(0) = PipeHz + 9
'            BasePt(1) = -PipeV + 12
'            BasePt(2) = 0
'            RecPipe(0) = BasePt(0) - PipeB / 1000 / 2
'            RecPipe(1) = BasePt(1) - PipeH / 1000 / 2
'            RecPipe(2) = BasePt(0) + PipeB / 1000 / 2
'            RecPipe(3) = BasePt(1) - PipeH / 1000 / 2
'            RecPipe(4) = BasePt(0) + PipeB / 1000 / 2
'            RecPipe(5) = BasePt(1) + PipeH / 1000 / 2
'            RecPipe(6) = BasePt(0) - PipeB / 1000 / 2
'            RecPipe(7) = BasePt(1) + PipeH / 1000 / 2
'            Set AcadRecPipe = AcadDoc.ModelSpace.AddLightWeightPolyline(RecPipe)
'            AcadRecPipe.Color = ColorNum
'            AcadRecPipe.Closed = True
'            AcadRecPipe.Update
'            Dimpt1(0) = BasePt(0)
'            Dimpt1(1) = Abs(BaseY) + 12
'            Dimpt2(0) = BasePt(0)
'            Dimpt2(1) = 12
'            Set DimLine = AcadDoc.ModelSpace.AddLine(Dimpt1, Dimpt2)
'            DimLine.LineType = "HIDDEN"
'            DimLine.Update
'            TextPt(0) = BasePt(0) - PipeB / 2000 - 0.3
'            TextPt(1) = BasePt(1)
'            Set AcadText = AcadDoc.ModelSpace.AddText(PipeType, TextPt, 0.2)
'            AcadText.Alignment = acAlignmentLeft
'            AcadText.Color = ColorNum
'            AcadText.Update
'        End If
End If

    AcadDoc.SaveAs ThisWorkbook.Path & "\" & ThisWorkbook.Worksheets("工作表1").Cells(k, 1).Value & ".dwg"
    AcadDoc.Application.ZoomExtents
    End If
Next

'斷面數量測試
''MsgBox (DatCount)
Z:
'        Set Currentplot = AcadDoc.Plot
'        AcadDoc.SetVariable BACKGROUNDPLOT, 0
'
'        AcadDoc.ActiveLayout.ConfigName = "DWG to PDF.pc3" ' Your plot device.
'        AcadDoc.ActiveLayout.CanonicalMediaName = "ISO A3 (420.00 x 297.00 公釐)"
'
'        AcadDoc.ActiveLayout.StandardScale = acScaleToFit
'        AcadDoc.Application.ZoomExtents
'        Currentplot.PlotToDevice
        AcadDoc.SendCommand "-plot" & vbCr & "Y" & vbCr & vbCr & "DWG To PDF.pc3" & vbCr & "ISO A3 (420.00 x 297.00 公釐)" & vbCr & "M" & vbCr & "L" & vbCr & "N" & vbCr & "E" & vbCr & "12.1=1" & vbCr & "C" & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr
'        AcadDoc.SendCommand "-plot" & vbCr & "Y" & vbCr & vbCr & "DWG To PDF.pc3" & vbCr & "w" & vbCr & "M" & vbCr & "L" & vbCr & "N" & vbCr & "w" & vbCr & "0.7615,0.7615" & vbCr & "41.1231,28.9285" & vbCr & "f" & vbCr & "C" & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr
        AcadDoc.Close
MsgBox ("Done")
End Sub
