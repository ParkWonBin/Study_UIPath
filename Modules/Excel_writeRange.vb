out_WriteRange = Function (strFileName As String, dtTemp As System.Data.DataTable, sheetName As String)
	Dim _excel As Microsoft.Office.Interop.Excel.Application
	Dim wBook As Microsoft.Office.Interop.Excel.Workbook
	Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
	Dim dt As System.Data.DataTable
	Dim dc As System.Data.DataColumn
	Dim dr As System.Data.DataRow
	Dim colIndex As Integer = 0
	Dim rowIndex As Integer = 0
	
	' 파일명 확인
	If Not strFileName.contains(":") ' 절대경로는 C: 포함하므로 : 로 구분
		strFileName = Environment.CurrentDirectory+ "\" + strFileName
	End If
	
	' 초기변수 설정
	_excel = New Microsoft.Office.Interop.Excel.Application
	wBook = _excel.Workbooks.Add()
	wSheet = CType(wBook.ActiveSheet,worksheet)
	' wSheet = CType(wBook.Worksheets.Add() ,Worksheet)
	wSheet.Name = sheetName
	
	' data 작성
	dt = dtTemp
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
	
	' 열 너비 설정
	wSheet.Columns.AutoFit()
	
	' 파일 존재여부 확인
	If System.IO.File.Exists(strFileName) Then
	    System.IO.File.Delete(strFileName)
	End If
	
	' 저장 및 종료
	wBook.SaveAs(strFileName)
	wBook.Close()
	_excel.Quit()
	Return "저장 성공 : "+vbnewline+strFileName
End Function
