Attribute VB_Name = "Initialize"
Sub openCAD()
Dim objAcadEntity As AcadEntity
Dim objAcadText As AcadText
Dim objAcadCircle As AcadCircle
Dim objAcadSelectionSet As AcadSelectionSet
Dim varPoint(1) As Variant
Dim varCenter As Variant

'On Error Resume Next
'�}��CAD�ɮ�
AcadFileName = InputBox("AutoCAD�ɮצW��")
Set AcadDoc = AcadApp.Documents.Open(ThisWorkbook.Path & "\" & AcadFileName)
AcadDoc.Application.Visible = True
AcadDoc.WindowState = acMax

ThisWorkbook.Worksheets("Initialize").Cells(1, 1).Value = "�ɮצW��:" & AcadFileName
ThisWorkbook.Worksheets("Initialize").Cells(2, 1).Value = Date
k = 4
q = 4
End Sub

