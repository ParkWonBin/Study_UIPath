'[ Static 변수 선언 ] =============================
' 인수로 주고받지 않아도 해당 객체에 바로 접근하기 위해 Static 변수로 뺴기
Static App_excel As Microsoft.Office.Interop.Excel.Application
Static wBook As Microsoft.Office.Interop.Excel.Workbook
Static wSheet As Microsoft.Office.Interop.Excel.Worksheet
'[ Sub Process 선언 ] ============================
Dim Act_Excel_Open As System.Action(Of String, String) = Sub(_FilePath As String, _SheetName As String)
	console.writeline("Open Excel")
	App_excel = New Microsoft.Office.Interop.Excel.Application
	App_excel.Visible = True
	App_excel.ScreenUpdating = True
	App_excel.DisplayAlerts=False
	'----------------------------------------
	console.writeline("Open Workbook : "+_FilePath)
	If Not String.isNullOrWhiteSpace(_FilePath) Then 
		wBook = App_excel.workbooks.Open( _ 
			FileName:=_FilePath, _ 
			UpdateLinks:= False, _
      [ReadOnly]:= False _
		)
	Else 
		wBook = CType(App_excel.Workbooks.Add(), Microsoft.Office.Interop.Excel.Workbook)
	End If
	'----------------------------------------
	Try
		wSheet = CType(wBook.Sheets(_SheetName), Microsoft.Office.Interop.Excel.Worksheet)
		console.writeline("Get WorkSheet : "+_SheetName)
	Catch e As System.Exception
		console.writeline("Add WorkSheet : "+_SheetName)
		wSheet = CType(wBook.Worksheets.Add() ,Microsoft.Office.Interop.Excel.Worksheet)
		wSheet.Name = _SheetName
	End Try
End Sub
'=============================
Dim Act_Excel_Close As System.Action(Of String) = Sub(_FilePath As String)
	If Not String.isNullOrWhiteSpace(_FilePath) Then 
		' 다른 이름으로 저장 설정
		If System.IO.File.Exists(_FilePath) Then
    		System.IO.File.Delete(_FilePath)
		End If	
		
		' 파일 서식 지정
		Dim _extention As Microsoft.Office.Interop.Excel.XlFileFormat		
		Select Case System.IO.Path.GetExtension(_FilePath).ToUpper
			Case Is = ".XLSX"
				'51 | .xlsx | Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook
				'61 | .xlsx | Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLStrictWorkbook
				_extention = Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLStrictWorkbook
			Case Is = ".CSV"
				'06 | .CSV | Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV
				'23 | .CSV | Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows
				_extention = Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV
			Case Is = ".XLSB"
				'40 | .xlsb | Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel12		
				_extention = Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel12	
			Case Is = ".XML"
				'46 | .xml | Microsoft.Office.Interop.Excel.XlFileFormat.xlXMLSpreadsheet
				_extention = Microsoft.Office.Interop.Excel.XlFileFormat.xlXMLSpreadsheet
			Case Is = ".XLSM"
				'52 | .xlsm| Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled
				_extention = Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled
			Case Else
				Dim _dirPath As String = System.IO.Path.GetDirectoryName(_FilePath)
				Dim _fileName As String = System.IO.Path.GetFileNameWithoutExtension(_FilePath)
				_FilePath = System.IO.Path.Combine(_dirPath, _fileName+".xlsx")
				_extention = Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLStrictWorkbook
		End Select
		console.writeline("다른 이름으로 저장 : "+_FilePath)
		wBook.SaveAs( _
			FileName:=_FilePath, _
			FileFormat := _extention, _
      Password := ""	_ 
    )
	Else 
		console.writeline("WorkBook 저장 ")
		wBook.Save()
	End If
	wBook.Close()
	'----------------------------------------
	App_excel.ScreenUpdating = True
	App_excel.DisplayAlerts=True
	App_excel.Quit()
	
	Threading.Thread.Sleep(2000)
End Sub
'====================================
'[ Main 시작 ] ============================
'인수 확인 및 Excel 열기
'Str_Wb_FilePath 비어있으면 새 파일 생성.
Dim Str_Wb_FilePath As String = ""
Dim Str_SheetName As String = "Config"
Act_Excel_Open(Str_Wb_FilePath, Str_SheetName)
'----------------------------------
console.writeline("Process")
App_excel.Cells(1, 1) = "Name"
App_excel.Cells(1, 2) = "Value"
App_excel.Cells(1, 3) = "Description"
wSheet.Range("A1:C1").Font.Bold = True 
wSheet.Range("A1:C1").Interior.Color = Color.LightGray
wSheet.Columns.AutoFit()

wSheet = CType(wBook.Sheets("Sheet1"), Microsoft.Office.Interop.Excel.Worksheet)
wSheet.Delete
'----------------------------------
' 엑실 종료 
'Str_Wb_FilePath 비어있으면 [저장], 안비어있으면 [다른 이름으로 저장]
Str_Wb_FilePath = System.IO.Path.Combine(System.Environment.CurrentDirectory, "Config.xlsx")
Act_Excel_Close(Str_Wb_FilePath)
