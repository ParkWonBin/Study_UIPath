'out_Bake_Config_AsExcel = Function ( dic_config As dictionary(Of String, String) )
	Dim excel As New Microsoft.Office.Interop.Excel.Application
	Dim wb As Microsoft.Office.Interop.Excel.Workbook
	Dim ws As Microsoft.Office.Interop.Excel.Worksheet
	Dim strFileName As String = Environment.CurrentDirectory+"\"+now.tostring("yyMMdd")+"_Bake_Config.xlsx"
		
	' 초기변수 설정
	wb = excel.Workbooks.Add()
	ws = CType(wb.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
	ws.Name = "Config"
	
	' Header 작성 Cells(row, col)
	 excel.Cells(1, 1) = "Name"
	 excel.Cells(1, 2) = "Value"
	 excel.Cells(1, 3) = "Description"
	 ws.Range("A1:C1").Font.Bold = True 
	 ws.Range("A1:C1").Interior.Color = Color.LightGray
	
	Dim rowIndex As Integer = 1
	Dim keys As String()
	keys = dic_config.keys().toarray
	system.array.sort(keys)
	For Each key As String In keys 
	    rowIndex = rowIndex + 1
		excel.Cells(rowIndex, 1) = key
	 	excel.Cells(rowIndex, 2) = dic_config(key)
	Next
	
	' 열 너비 설정
	ws.Columns.AutoFit()
	
	' 파일 존재여부 확인
	If System.IO.File.Exists(strFileName) Then
	    System.IO.File.Delete(strFileName)
	End If
	
	' 저장 및 종료
	wb.SaveAs(strFileName)
	wb.Close()
	excel.Quit()
	'Return "저장 성공 : "+vbnewline+strFileName
'End Function
