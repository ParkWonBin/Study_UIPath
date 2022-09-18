Dim Str_FilePath As String = System.IO.Path.Combine(System.Environment.CurrentDirectory,"test.xlsx")
Dim Str_SheetName As String = "New Sheet"

Dim App_excel As New Microsoft.Office.Interop.Excel.Application
App_excel.Visible = False
App_excel.DisplayAlerts=False

' 새 워크북 생성
Dim wb As Microsoft.Office.Interop.Excel.Workbook = App_excel.Workbooks.Add()
Dim ws As Microsoft.Office.Interop.Excel.Worksheet = CType(wb.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet) 
	
With ws
	.Name = Str_SheetName
	'Data 설정, Offset. 1부터 시작
	With .Range("A1")
		.Cells(1,1) = "Name"
		.Cells(1,2) = "Value"
		.Cells(1,3) = "Description"
	End With
	'Data서식 설정
	With .Range("A1:C1")
		.Font.Bold = True 
		.Interior.Color = System.Drawing.Color.LightGray
		.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
		.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
	End With
	.Columns.AutoFit()
End With

wb.SaveAs(Str_FilePath)
wb.Close()
App_excel.Quit()
