Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open _
    ("C:\Users\kabhishek\Desktop\TestCases.xls")
Set workSheet =  objExcel.ActiveWorkbook.Worksheets("Module")	
Set Cells = workSheet.Cells
str = Cells(1,2).Value
MsgBox str 