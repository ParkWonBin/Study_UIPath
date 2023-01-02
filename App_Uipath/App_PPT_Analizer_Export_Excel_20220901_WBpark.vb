'2022.09.01. 박원빈 프로
'현제 열려있는 PPT파일의 Data를 Excel로 변환하는 모듈입니다.
'WBpark__App_PPT_Analizer_Export_Excel_20220901
'--------------------
Dim Str_Excel_FilePath As String = in_Str_Excel_FilePath
Dim Str_PPT_FilePath As String = in_Str_PPT_FilePath

If String.IsNullOrWhiteSpace(Str_Excel_FilePath) Then
	Str_Excel_FilePath = path.Combine(Environment.CurrentDirectory,"PPT_Analizer.xlsx")
End If 
'----------------------------------------------------------------------------------
Dim Dic_PPT As New system.Collections.Generic.Dictionary(Of String, System.Data.DataTable)
Dim StrArr_ColNames As String() = "SlideIndex | Slide Name | Shape Id | Type | Shape Name | ZOrderPosition | HasTable[Row,Col] | Top | Left | Width | Height | Text | Remarks".split("|"c).Select(Function(x) x.Trim).ToArray
Dim Dic_ShapeType As New Dictionary(Of Integer,String) From {
{-2,"msoShapeTypeMixed"},
{1,"msoAutoShape"},
{2,"msoCallout"},
{3,"msoChart"},
{4,"msoComment"},
{5,"msoFreeform"},
{6,"msoGroup"},
{7,"msoEmbeddedOLEObject"},
{8,"msoFormControl"},
{9,"msoLine"},
{10,"msoLinkedOLEObject"},
{11,"msoLinkedPicture"},
{12,"msoOLEControlObject"},
{13,"msoPicture"},
{14,"msoPlaceholder"},
{15,"msoTextEffect"},
{16,"msoMedia"},
{17,"msoTextBox"},
{18,"msoScriptAnchor	"},
{19,"msoTable"},
{20,"msoCanvas"},
{21,"msoDiagram"},
{22,"msoInk"},
{23,"msoInkComment"},
{24,"msoIgxGraphic"},
{25,"msoSlicer"},
{26,"msoWebVideo"},
{27,"msoContentApp"},
{28,"msoGraphic"},
{29,"msoLinkedGraphic"},
{30,"mso3DModel"},
{31,"msoLinked3DModel"}
}
'----------------------------------------------------------------------------------
'PPT 초기 세팅
Dim App_ppt As New Microsoft.Office.Interop.PowerPoint.Application
Dim ppt_source As Microsoft.Office.Interop.PowerPoint.Presentation
If String.IsNullOrWhiteSpace(Str_PPT_FilePath) Then
	' 열려있는 PPT 읽기
	For Each ppt_tmp As Microsoft.Office.Interop.PowerPoint.Presentation In App_ppt.Presentations
	   ppt_source = ppt_tmp
	Next ppt_tmp
Else 
	'경로에서 PPT 열기
	ppt_source = App_ppt.Presentations.Open( FileName:=Str_PPT_FilePath, _
		ReadOnly:=Microsoft.Office.Core.MsoTriState.msoTrue, _
		Untitled:=Microsoft.Office.Core.MsoTriState.msoFalse, _
		WithWindow:=Microsoft.Office.Core.MsoTriState.msoFalse
	) ' Untitled 속성을 True로 하면, 파일을 복사해서 새로 여는 것과 같습니다.
End If 
'----------------------------------------------------------------------------------
'Data 추출 시작
console.writeline("PPT 데이터 추출 시도")
For Each Sld_tmp As Microsoft.Office.Interop.PowerPoint.Slide In ppt_source.Slides
	
	'해당 슬라이드에 대한 DT 생성
	console.writeLine("데이터 추출 : 슬라이드 "+Sld_tmp.SlideIndex.ToString)
	Dim DT_Slide As New System.Data.DataTable
	For Each ColName As String In StrArr_ColNames
		DT_Slide.Columns.Add(ColName, GetType(String))
	Next

	' 슬라이드 Attach
    With Sld_tmp
        For Each Shp_tmp As Microsoft.Office.Interop.PowerPoint.Shape In .Shapes
			
			' 해당 슬라이드 DT 행추가
			Dim NewRow As System.Data.DataRow = DT_Slide.Rows.Add()
			
			With Shp_tmp
				NewRow("SlideIndex") = CStr(CType( .parent, Microsoft.Office.Interop.PowerPoint.Slide).SlideIndex )
				NewRow("Slide Name") = CStr(CType( .parent, Microsoft.Office.Interop.PowerPoint.Slide).Name )
				NewRow("Shape Id") = CStr(.Id)
				NewRow("Shape Name") = CStr(.Name)
				NewRow("ZOrderPosition") = CStr(.ZOrderPosition)
				NewRow("HasTable[Row,Col]") = If(CBool(.HasTable),"[O]","[X]" )
				NewRow("Top") = CStr(.Top)
				NewRow("Left") = CStr(.Left)
				NewRow("Width") = CStr(.Width)
				NewRow("Height") = CStr(.Height)
				NewRow("Text") = If( CBool(.HasTextFrame), CStr(.TextFrame.TextRange.Text), "")
				NewRow("Type") = If( Dic_ShapeType.containskey(CInt(.Type)), Dic_ShapeType(CInt(.Type)), CStr(.Type))
				NewRow("Remarks") = ""
			End With
			
			' Shape가 Table이 아니면 다음 항목으로 -----------------
			If Not CBool(Shp_tmp.HasTable) Then
				Continue For
			End If
			
			' Shape가 Table이면 해당 Table 내용도 기록
			For int_i As Integer = 1 To Shp_tmp.Table.Rows.Count
				For int_j As Integer = 1 To Shp_tmp.Table.Columns.Count
					With Shp_tmp.Table.Rows(int_i).Cells(int_j).Shape
						Dim NewTableRow As System.Data.DataRow = DT_Slide.Rows.Add()
						NewTableRow("SlideIndex") = NewRow("SlideIndex").ToString
						NewTableRow("Slide Name") = NewRow("Slide Name").ToString
						NewTableRow("Shape Id") = NewRow("Shape Id").ToString
						NewTableRow("Shape Name") = NewRow("Shape Name").ToString
						NewTableRow("HasTable[Row,Col]") = String.format("{2} [{0},{1}] ", int_i.ToString("00"), int_j.ToString("00"), NewRow("Shape Name").ToString)
						NewTableRow("Top") = CStr(.Top)
						NewTableRow("Left") = CStr(.Left)
						NewTableRow("Width") = CStr(.Width)
						NewTableRow("Height") = CStr(.Height)
						NewTableRow("Text") = CStr(.TextFrame.TextRange.Text)
						NewTableRow("Type") = "Table Data"
						NewTableRow("Remarks") = ""
					End With
				Next int_j
			Next int_i
        Next Shp_tmp
    End With
	
	'DIctionary에 해당 슬라이드 정보 등록
	Dic_PPT("#"+Sld_tmp.SlideIndex.ToString) = _ 
		DT_Slide.AsEnumerable _ 
			.OrderBy(Function(row) row.item("HasTable[Row,Col]").tostring) _ 
			.Thenby(Function(row) CDbl(row.item("Top")) ) _ 
			.Thenby(Function(row) CDbl(row.item("Left")) ) _ 
		.copytoDataTable
Next Sld_tmp
'----------------------------------------------------------------------------------
'PPT 종료
'If ppt_source IsNot Nothing Then
    'ppt_source.Close()
'End If
'App_ppt.Quit
'----------------------------------------------------------------------------------
out_Dic_PPT = Dic_PPT
'----------------------------------------------------------------------------------
'Excel 파일로 작성
console.writeline("Excel App 열기 시도")

' Init01 - Open Excel
Dim App_excel As New Microsoft.Office.Interop.Excel.Application
App_excel.Visible = False
App_excel.ScreenUpdating = True
App_excel.DisplayAlerts = False

Dim wBook As Microsoft.Office.Interop.Excel.Workbook = App_excel.Workbooks.add()
Dim wSheet As Microsoft.Office.Interop.Excel.WorkSheet
'---------------------------------------------------------------------------------
For Each Str_SheetName As String In Dic_PPT.Keys.Reverse
	'Get DT
	Dim DT_tmp As System.Data.DataTable = Dic_PPT(Str_SheetName)
	Dim int_col_Cnt As Integer = DT_tmp.Columns.Count
	Dim int_rows_Cnt As Integer =DT_tmp.rows.Count
	Console.WriteLine(Str_SheetName)
	
	'Add Sheet
	 wSheet = CType(wBook.Worksheets.Add(), Microsoft.Office.Interop.Excel.Worksheet)
	wSheet.Name = Str_SheetName
	
	'Add Header
	For int_col_idx As Integer = 0 To int_col_Cnt-1
		wSheet.Cells(1,1+int_col_idx) = DT_tmp.Columns(int_col_idx).ColumnName
	Next int_col_idx
	wSheet.Range(wSheet.Cells(1,1),wSheet.Cells(1,int_col_Cnt)).Font.Bold = True
	wSheet.Range(wSheet.Cells(1,1),wSheet.Cells(1,int_col_Cnt)).Interior.Color = Color.LightGray
	
	'Write Range
	For int_row_idx As Integer = 0 To int_rows_Cnt-1
		For int_col_idx As Integer = 0 To StrArr_ColNames.Count-1
			wSheet.Cells(2+int_row_idx,1+int_col_idx) = DT_tmp.Rows(int_row_idx).Item(int_col_idx).ToString.Trim
		Next int_col_idx
	Next int_row_idx
	
	'Auto Fit Columns
	wSheet.columns.AutoFit()
Next Str_SheetName
'----------------------------------------------------------------------------------
'Delete Default Sheet
wSheet = CType(wBook.Sheets("Sheet1"), Microsoft.Office.Interop.Excel.Worksheet)
wSheet.Delete()
'Delete ResultFilePath if File Exist
System.IO.File.Delete(Str_Excel_FilePath)
'----------------------------------------------------------------------------------
'End01 - Save WorkBook
wBook.SaveAs(Str_Excel_FilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLStrictWorkbook)
wBook.Close()
'End02 - Close App
App_excel.ScreenUpdating = True
App_excel.DisplayAlerts = True
App_excel.Quit()
'End03 - Wait DRM close
Threading.Thread.Sleep(2000)
'exception.InnerException.Message
'exception.InnerException.TargetSite
'exception.InnerException.Source
'exception.InnerException.StackTrace
'exception.tostring