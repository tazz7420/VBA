Attribute VB_Name = "Z_公用函數"
Public Function Pi() As Double
  Pi = 4# * Atn(1#)
End Function

'
'徑度值 > (2 * pi) 則 - (2 * pi) ; 徑度值 < 0 則 + (2 * pi)
'
Public Function Ang2Pi(ByVal rad As Double) As Double
  If rad >= (2# * Pi) Then
     rad = rad - (2# * Pi)
  ElseIf rad < 0# Then
     rad = rad + (2# * Pi)
  Else
  End If
  Ang2Pi = rad
End Function

'方位角 轉換為 AutoCAD 角度系統
Public Function AzToAcadAngle(ByVal tAz As Double) As Double
   AzToAcadAngle = Ang2Pi((Pi * 2.5) - tAz)
End Function


'
'求取兩點坐標之距離值
'
Public Function Hdist(ByVal Iyn As Double, ByVal Ixe As Double, ByVal Byn As Double, ByVal Bxe As Double) As Double
  Hdist = Sqr(((Iyn - Byn) * (Iyn - Byn)) + ((Ixe - Bxe) * (Ixe - Bxe)))     '傳回值
End Function
'
'求取第一點坐標至第二點坐標之方位角值
'
Public Function Pol(ByVal Iyn As Double, ByVal Ixe As Double, ByVal Byn As Double, ByVal Bxe As Double) As Double
  Dim A As Double, B As Double, C As Double
  A = Byn - Iyn
  B = Bxe - Ixe
  If A = 0# And B = 0# Then
     'MsgBox "座標反算方位角錯誤，原因為二點座標相同，請檢查!", vbCritical + vbOKOnly
     '錯誤及警告訊息.Show
     '錯誤及警告訊息.TxtError.Text = 錯誤及警告訊息.TxtError.Text + "座標相同，反算方位角錯誤 , " & _
     '                               " Y= " & Format(iyn, "0.000") & "  X= " & Format(ixe, "0.000") & vbCrLf
     Pol = -9#
  ElseIf A = 0# And B > 0# Then
     Pol = Pi / 2#
  ElseIf A = 0# And B < 0# Then
     Pol = Pi * 1.5
  ElseIf A > 0# And B >= 0# Then
     Pol = Atn(B / A)
  ElseIf A > 0# And B < 0# Then
     Pol = Atn(B / A) + Pi * 2#
  ElseIf A < 0# Then
     Pol = Atn(B / A) + Pi
  Else
  End If
  
End Function

Public Function RbtSelectCrossing(ByRef tSelxxx As AcadSelectionSet, ByVal Xyz1 As Variant, ByVal Xyz2 As Variant, _
                             ByVal FilterGroupCount As Long, ParamArray FilterData()) As Boolean
   Dim i As Long, j As Long, L As Long, u As Long
   Dim tCode() As Integer
   Dim tData() As Variant
   
   ReDim tCode(0 To FilterGroupCount - 1): ReDim tData(0 To FilterGroupCount - 1)
   L = LBound(FilterData): u = UBound(FilterData)
   For i = L To u Step FilterGroupCount * 2
      For j = 0 To FilterGroupCount - 1
         tCode(j) = FilterData(i + (j * 2))
         tData(j) = FilterData(i + (j * 2) + 1)
      Next
      tSelxxx.Select acSelectionSetCrossing, Xyz1, Xyz2, tCode, tData
   Next
End Function

Public Function RbtAddNewSelection(tSelection As AcadSelectionSet, ByVal tSelectionName As String) As Boolean
   On Error Resume Next
   Set tSelection = ActiveDocument.SelectionSets.Add(tSelectionName)
   If Err Then
      Err.Clear
      ActiveDocument.SelectionSets(tSelectionName).Delete
      Set tSelection = ActiveDocument.SelectionSets.Add(tSelectionName)
   End If
End Function


