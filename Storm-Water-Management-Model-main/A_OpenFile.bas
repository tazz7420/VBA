Attribute VB_Name = "A_OpenFile"
Public AcadApp As New AcadApplication
Public AcadApp1 As New AcadApplication
Public AcadDoc As AcadDocument
Public AcadDoc1 As AcadDocument
Public M_Sheet(1 To 9) As Excel.Worksheet
Sub open_file()
Set AcadDoc = AcadApp.Documents.Open("C:\Users\tazz4\OneDrive\орн▒\20200829\Drawing6.dwg")
AcadDoc.Application.Visible = True
AcadDoc.WindowState = acMax

Set AcadDoc1 = AcadApp1.Documents.Open("C:\Users\tazz4\OneDrive\орн▒\20200829\Drawing5.dwg")
AcadDoc1.Application.Visible = True
AcadDoc1.WindowState = acMax

Set M_Sheet(1) = ThisWorkbook.Worksheets("JUNCTIONS")
Set M_Sheet(2) = ThisWorkbook.Worksheets("CONDUITS")
Set M_Sheet(3) = ThisWorkbook.Worksheets("XSECTIONS")
Set M_Sheet(4) = ThisWorkbook.Worksheets("TRANSECTS")
Set M_Sheet(5) = ThisWorkbook.Worksheets("COORDINATES")
Set M_Sheet(6) = ThisWorkbook.Worksheets("VERTICES")
Set M_Sheet(7) = ThisWorkbook.Worksheets("Sheet7")
Set M_Sheet(8) = ThisWorkbook.Worksheets("Sheet1")
Set M_Sheet(9) = ThisWorkbook.Worksheets("Sheet2")

'Call DrawLink
'Call MissingTransects
Call FindDhTf
'Call CheckTransects
Call FindMaxDepth
Call JunctionsMaxDepth
Call FixConduits
'Call FindTransects

End Sub

