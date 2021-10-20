out_WriteRange = Function (strFileName As String, dtTemp As System.Data.DataTable, sheetName As String)
	Dim _excel As Microsoft.Office.Interop.Excel.Application
	Dim wBook As Microsoft.Office.Interop.Excel.Workbook
	Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
	Dim dt As System.Data.DataTable
	Dim dc As System.Data.DataColumn
	Dim dr As System.Data.DataRow
	Dim colIndex As Integer = 0
	Dim rowIndex As Integer = 0
	
	If Not strFileName.contains(":") ' 절대경로는 C: 포함하므로 : 로 구분
		strFileName = Environment.CurrentDirectory+ "\" + strFileName
	End If
	
	dt = dtTemp
	_excel = New Microsoft.Office.Interop.Excel.Application
	wBook = _excel.Workbooks.Add()
	' wSheet = CType(wBook.Worksheets.Add() ,Worksheet)
	wSheet = CType(wBook.ActiveSheet,worksheet)
	wSheet.Name = sheetName
	
	For Each dc In dt.Columns
	    colIndex = colIndex + 1
	    _excel.Cells(1, colIndex) = dc.ColumnName
	Next
	For Each dr In dt.Rows
	    rowIndex = rowIndex + 1
	    colIndex = 0
	    For Each dc In dt.Columns
	        colIndex = colIndex + 1
	        _excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
	    Next
	Next
	wSheet.Columns.AutoFit()
	If System.IO.File.Exists(strFileName) Then
	    System.IO.File.Delete(strFileName)
	End If
	' console.WriteLine("5")
	wBook.SaveAs(strFileName)
	wBook.Close()
	_excel.Quit()
	Return "저장 성공 : "+vbnewline+strFileName
End Function
