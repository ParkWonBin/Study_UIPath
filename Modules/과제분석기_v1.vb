'WBpark_20220128_App_Bake_Selector_CSV_From_ProjectXaml
'Dim in_Str_ProjDir  As String = ""
'Dim in_Str_OutputDir As String = "C:\Users\H2109941\Desktop\Report"
'Dim out_Dt_Project_Activity_info As System.Data.DataTable
'------------------------------------------------------------------------------------------------------------------------------------------
'Fnc_Get_All_ProjDir
Dim Fnc_Get_All_ProjDir As System.Func(Of String, String()) = Function(str_path As String) As String()
	Dim list_str_dir As New System.Collections.Generic.List(Of String) 
	Dim list_str_proj As New System.Collections.Generic.List(Of String)
	
	list_str_dir.Add(str_path)
	While list_str_dir.Count <> 0
		str_path = list_str_dir.Last
		list_str_dir.RemoveAt(list_str_dir.Count -1)
		If System.IO.Directory.GetFiles(str_path).Where(Function(x) x.ToUpper.Contains(".XAML")).Count > 0 Then
			list_str_proj.Add(str_path)
		Else
			list_str_dir.AddRange(System.IO.Directory.GetDirectories(str_path).ToArray)	
		End If
	End While
	Return list_str_proj.ToArray
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
'Fnc_Get_All_XamlFiles
Dim Fnc_Get_All_XamlFiles As System.Func(Of String, String()) = Function(str_path As String) As String()
	Dim list_str_dir As New System.Collections.Generic.List(Of String) 
	Dim list_str_file As New System.Collections.Generic.List(Of String)
	
	list_str_dir.Add(str_path)
	While list_str_dir.Count <> 0
		str_path = list_str_dir.Last
		list_str_dir.RemoveAt(list_str_dir.Count -1)
		list_str_dir.AddRange(System.IO.Directory.GetDirectories(str_path).Where(Function(x) Not x.Contains(".screenshots")).ToArray)
		list_str_file.AddRange(System.IO.Directory.GetFiles(str_path))
	End While
	
	Dim StrArr_Files As String() = list_str_file.Where(Function(x) x.Contains(".xaml")).ToArray
	'Console.WriteLine(String.Format(" 파일 : {0} 개 확인", StrArr_Files.Count.ToString))	
	Return StrArr_Files
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
'Fnc_Xnode_Get_Parent
Dim Fnc_Xnode_Get_Parent As System.Func(Of System.Xml.XmlNode, Integer,  System.Xml.XmlNode) = Function(Node As System.Xml.XmlNode,  P_Depth As Integer) As System.Xml.XmlNode
	For i As Integer = 1 To P_Depth
		If Node.ParentNode.Name <> "Activity"  Then ' 최상위 노드에서 parent 호출시 에러 발생. 에러 회피용
			Node =  Node.ParentNode
		End If
	Next 
	Return Node
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
'Fnc_Xnode_Get_Attribute
Dim Fnc_Xnode_Get_Attribute As System.Func(Of System.Xml.XmlNode, String, String)  = Function(Node As System.Xml.XmlNode, AttrName As String) As String
	If Node.Attributes.count > 0 Then 
		If Enumerable.Range(0, Node.Attributes.count).Select(Function(i) Node.Attributes.ItemOf(i).Name.ToString ).ToArray.Contains(AttrName) Then
			Return Node.Attributes(AttrName).value.ToString 
		End If
    End If
	Return ""	
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
'Fnc_Xnode_Get_Attributes
Dim Fnc_Xnode_Get_Attributes As System.Func(Of System.Xml.XmlNode, String(), String)  = Function(Node As System.Xml.XmlNode, AttrNames As String() ) As String
    If Node.Attributes.count = 0 Then
        Return ""
    End If
	
    Dim Node_Attr As String() = Enumerable.Range(0, Node.Attributes.count).Select(Function(x,i) Node.Attributes.ItemOf(i).Name.ToString).where(Function(x) AttrNames.Contains(x) ).Select(Function(x) String.Format("{0} : {1} ,",x, Node.Attributes(x).value.ToString) ).ToArray
    Return Join(Node_Attr, chr(10).ToString )
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
'Fnc_Build_DT
Dim Fnc_Build_DT As System.Func(Of String(), System.Data.DataTable) = Function( StrArr_ColNames As String() ) As System.Data.Datatable
	Dim Dt_tmp As New DataTable
	For Each colName As String In StrArr_ColNames
            Dt_tmp.Columns.Add(colName, System.Type.GetType("System.String"))
    Next
	Return Dt_tmp
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
'Fnc_Extract_Selector_From_Xaml
Dim Fnc_Extract_Selector_From_Xaml As System.Func(Of String, System.Data.DataTable, System.Data.DataTable) = Function( XamlPath As String, Dt_result As System.Data.DataTable) As System.Data.DataTable
    ' dt 확인
    If Dt_result Is Nothing OrElse Dt_result.Columns.count = 0 Then
        Dt_result = Fnc_Build_DT("is_Web_UI|DirectoryName|FileName|ActivityName|DisplayName|Refid|Selector|Selector_Recommended|ReMarks".split("|"c))
    End If

    ' Xaml 읽기 시도.
    Dim doc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
    'console.WriteLine("Xaml 읽기 시도 : "+vbNewLine+XamlPath)
    doc.Load(XamlPath)

    ' 데이터 추출 시도
    Dim Str_DirName As String = (New System.IO.FileInfo(XamlPath)).DirectoryName.Split("\"c).last
    Dim Str_FileName As String = (New System.IO.FileInfo(XamlPath)).Name

    ' 전체 요소 선택 > Attribute 있는 요소만 선택 > Selector 있는 요소만 선택
    Dim tags As System.Xml.XmlNodeList = doc.GetElementsByTagName("*")
    Dim Arr_Nodes As System.Xml.XmlNode() = Enumerable.range(0, tags.Count).Select(Function(i) tags(i) ).where(Function(x) Not String.IsNullOrWhiteSpace(Fnc_Xnode_Get_Attribute(x, "Selector"))).ToArray
    
    ' 데이터 추출 및 DT에 갱신
    If Arr_Nodes.count > 0 
		
        ' GetValue, SetValue, TypeInto 등 처리 | 0계층 위
		Dim StrArr_Refid As String() = Arr_Nodes.Select(Function(x) Fnc_Xnode_Get_Attribute( x, "sap2010:WorkflowViewState.IdRef") ).toArray
		Dim StrArr_ReMarks As String() = Enumerable.Repeat(Of String)(" ", Arr_Nodes.count).ToArray
		Dim StrArr_ActivityNames As String() = Arr_Nodes.Select(Function(x) x.Name ).toArray
		Dim StrArr_DisplayNames As String() = Arr_Nodes.Select(Function(x) Fnc_Xnode_Get_Attribute( x, "DisplayName") ).toArray
		Dim StrArr_Selectors As String() = Arr_Nodes.Select(Function(x) x.Attributes("Selector").value.Replace(chr(10).ToString,"").Replace(chr(13).ToString,"") ).toArray
		Dim StrArr_Selectors_RCMD As String() = StrArr_Selectors.Select(Function(x) x.Replace("'iexplore.exe'","'msedge.exe'") ).toArray
		Dim StrArr_Is_Web_UI As String() = StrArr_Selectors.Select(Function(x) (x.Contains("<html") OrElse x.Contains("<webctrl") OrElse x.Contains("iexplore") OrElse x.Contains("x:Null") ).ToString ).toArray
		
		' WaitUiElementVanish, Attach 등 처리 | 2계층 위
		StrArr_Refid = StrArr_ActivityNames.Select(Function(x,i) If( x <> "ui:Target" ,  StrArr_Refid(i) , Fnc_Xnode_Get_Attribute( Fnc_Xnode_Get_Parent(Arr_Nodes(i),2), "sap2010:WorkflowViewState.IdRef")  ) ).toArray
		StrArr_ReMarks = StrArr_ActivityNames.Select(Function(x,i) If( x <> "ui:Target",  StrArr_ReMarks(i) , Fnc_Xnode_Get_Attributes(Fnc_Xnode_Get_Parent(Arr_Nodes(i),2), {"BrowserType","UiBrowser","Exists","SendWindowMessages","SimulateClick","SimulateType"}))).toArray
        StrArr_DisplayNames = StrArr_ActivityNames.Select(Function(x,i) If(  x <> "ui:Target",  StrArr_DisplayNames(i) , Fnc_Xnode_Get_Attribute( Fnc_Xnode_Get_Parent(Arr_Nodes(i),2), "DisplayName")  ) ).toArray
        StrArr_ActivityNames = StrArr_ActivityNames.Select(Function(x,i) If(  x <> "ui:Target", x , Fnc_Xnode_Get_Parent(Arr_Nodes(i),2).Name ) ).toArray
		
		' Open, Attach, Browse scope 등 처리 | 6계층 위
		' 오래된 Xaml중에 Refid는 존재하나 DisplayName이 없는 파일들이 있다.
        StrArr_Refid = StrArr_ActivityNames.Select(Function(x,i) If( Not String.IsNullOrWhiteSpace(StrArr_DisplayNames(i)) AndAlso x.Contains("ui:"), StrArr_Refid(i), Fnc_Xnode_Get_Attribute( Fnc_Xnode_Get_Parent(Arr_Nodes(i), 6) ,"sap2010:WorkflowViewState.IdRef") ) ).toArray
		StrArr_ReMarks = StrArr_ActivityNames.Select(Function(x,i) If( Not String.IsNullOrWhiteSpace(StrArr_DisplayNames(i)) AndAlso x.Contains("ui:"), StrArr_ReMarks(i), Fnc_Xnode_Get_Attributes(Fnc_Xnode_Get_Parent(Arr_Nodes(i),6), {"BrowserType","UiBrowser","Url"}))).toArray
		StrArr_ActivityNames = StrArr_ActivityNames.Select(Function(x,i) If( Not String.IsNullOrWhiteSpace(StrArr_DisplayNames(i)) AndAlso x.Contains("ui:") , x , Fnc_Xnode_Get_Parent(Arr_Nodes(i) ,6).Name ) ).toArray
		StrArr_DisplayNames = StrArr_ActivityNames.Select(Function(x,i) If( Not String.IsNullOrWhiteSpace(StrArr_DisplayNames(i)) AndAlso x.Contains("ui:"), StrArr_DisplayNames(i) , Fnc_Xnode_Get_Attribute( Fnc_Xnode_Get_Parent(Arr_Nodes(i), 6) ,"DisplayName") ) ).toArray		
        For Each i As Integer In Enumerable.Range(0,StrArr_Selectors.Count)
                Dt_result.Rows.Add({StrArr_Is_Web_UI(i),Str_DirName, Str_FileName, StrArr_ActivityNames(i), StrArr_DisplayNames(i),StrArr_Refid(i), StrArr_Selectors(i),StrArr_Selectors_RCMD(i), StrArr_ReMarks(i)})
        Next
    End If
    Return Dt_result
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
'Fnc_fit_CSV_Cell
Dim Fnc_fit_CSV_Cell As System.Func(Of String, String) = Function(Str_source As String) As String
	Return Str_source.Replace(",","|").Replace(chr(10).ToString,"").Replace(chr(13).ToString,"")
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
'Convert_DT_to_CSV
Dim Convert_DT_to_CSV As System.Func(Of System.Data.DataTable, String) = Function(DT_Source As System.Data.DataTable) As String
	Dim colSep As String =","
	Dim rowSep As String = chr(13).ToString+chr(10).ToString
	Dim StrArr_cols As String() = Enumerable.Range(0,DT_Source.Columns.Count).Select(Function(x) DT_Source.Columns.Item(x).ColumnName).ToArray 
	Dim StrArr_rows As String()  = DT_Source.AsEnumerable.Select(Function(row) Join(row.ItemArray.Select(Function(x) Fnc_fit_CSV_Cell(x.ToString) ).toArray, colSep) ).ToArray
	Dim Str_CSV As String  = Join(StrArr_cols, colSep) & rowSep & Join(StrArr_rows, rowSep)
	Return Str_CSV 
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
'Write CSV
Dim Write_CSV As System.Func(Of String, String, String) = Function(Str_FilePath As String, Str_Contents As String) As String
	If Str_FilePath.Split("."c).Last.ToString.ToUpper <> "CSV" Then
		Str_FilePath = Str_FilePath + ".csv"
	End If
	If System.IO.File.Exists(Str_FilePath ) Then
		System.IO.File.Delete(Str_FilePath) 
	End If
	System.IO.File.WriteAllText(Str_FilePath,Str_Contents, System.Text.Encoding.UTF8)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
'Fnc_GetProperty
Dim Fnc_GetProperty As System.Func(Of System.Text.Json.JsonElement, String, System.Text.Json.JsonElement) = Function(Json_Sorce As System.Text.Json.JsonElement, Str_PropertyName As String) As System.Text.Json.JsonElement
	Try 
		Return Json_Sorce.GetProperty(Str_PropertyName)
	Catch ex As Exception
	End Try
		Return Json_Sorce
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
'Fnc_Extract_Proj_info
Dim Fnc_Extract_Proj_info As System.Func(Of String, String()) = Function(Str_Projpath As String) As String()
	' Return : ProjectName | StudioVersion | AutomaionVersion | Dependencies
	Dim StrArr_ProjExist As String() = System.IO.Directory.GetFiles(Str_Projpath).Where(Function(x) x.Contains("project.json") ).ToArray
	
	If StrArr_ProjExist.Count = 0 Then
		Return New String() {"","","",""}
	Else 
		Dim Str_Proj_JsonPath As String = StrArr_ProjExist.First.ToString
		Dim Json_Proj As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse( System.IO.File.ReadAllText(Str_Proj_JsonPath) )
		Dim Json_root As System.Text.Json.JsonElement = Json_Proj.RootElement
		Return New String(){  
			Fnc_GetProperty(Json_root,"name").Tostring,
			Fnc_GetProperty(Json_root,"studioVersion").Tostring ,
			Fnc_GetProperty(Fnc_GetProperty(Json_root,"dependencies"),"UiPath.UIAutomation.Activities").Tostring,
			Fnc_GetProperty(Json_root,"dependencies").Tostring 
			}	
	End If
End Function
'------------------------------------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------------------------------------
'main
Dim StrArr_ProjDir As String() = Fnc_Get_All_ProjDir(in_Str_ProjDir)
Dim Str_OutputDir As String = in_Str_OutputDir

If Not System.IO.Directory.Exists(System.IO.Path.Combine(Str_OutputDir, "Selector")) Then
	System.IO.Directory.CreateDirectory(System.IO.Path.Combine(Str_OutputDir, "Selector"))
End If	

Dim Dt_Project_Activity_info As System.Data.DataTable = Fnc_Build_DT("ProjName|StudioVersion|UIAutomaion.Ver|Total_Xaml_Count|Total_UI_Count|Web_Ui_Count|ElementExist_Count|GruopBy_FileName|GruopBy_Activites|Dependencies".Split("|"c))

Dim int_i As Integer = 0
For Each Str_ProjectPath As String In StrArr_ProjDir
	'Loop init Settings
	int_i += 1
	Dim Dtm_StartTime As System.DateTime = Now
	Dim Str_ProjectName As String = system.IO.Path.GetFileNameWithoutExtension(Str_ProjectPath)	
	console.WriteLine(String.Format("{0} / {1} | 추출시도 : {2}", int_i.ToString("000"), StrArr_ProjDir.Count.ToString("000"), Str_ProjectName) )
	Try 
		'Project_Version_info
		Dim StrArr_Proj_info As String() = Fnc_Extract_Proj_info(Str_ProjectPath)
		If String.IsNullOrWhiteSpace( StrArr_Proj_info.First ) Then
			StrArr_Proj_info(0) = Str_ProjectName
		End If
		'ProjectName | StudioVersion | UIAutomationVersion | Dependencies
		
		'Project_Selector_info
		Dim Dt_Project_Selector_info As New System.Data.DataTable
		Dim StrArr_XmalFiles As String() = Fnc_Get_All_XamlFiles(Str_ProjectPath)
		console.WriteLine(String.Format("{0} / {1} | 파일확인 : {2} 개", int_i.ToString("000"), StrArr_ProjDir.Count.ToString("000"), StrArr_XmalFiles.Count.ToString) )
		For Each xamlPath As String In StrArr_XmalFiles 
		    Dt_Project_Selector_info = Fnc_Extract_Selector_From_Xaml(xamlPath,Dt_Project_Selector_info)
			'is_Web_UI | DirectoryName | FileName | ActivityName | DisplayName | Refid | Selector | Selector_Recommended | ReMarks
		Next
		Write_CSV(System.IO.Path.Combine(Str_OutputDir, "Selector",Str_ProjectName), Convert_DT_to_CSV(Dt_Project_Selector_info) ) 
		
		'Project_Activity_info
		Dim StrArr_NewRow As String() = New String() {
			StrArr_Proj_info(0),
			StrArr_Proj_info(1),
			StrArr_Proj_info(2),
			StrArr_XmalFiles.Count.ToString,
			Dt_Project_Selector_info.Rows.Count.ToString,
			Dt_Project_Selector_info.AsEnumerable.Where(Function(row) row.Item(0).ToString.ToUpper = "TRUE" ).Count.ToString,
			Dt_Project_Selector_info.AsEnumerable.Where(Function(row) row.Item(3).ToString = "ui:UiElementExists" ).Count.ToString,
			Join(Dt_Project_Selector_info.AsEnumerable.GroupBy(Function(row) row("FileName").ToString).Select(Function(gp) String.Format("{0} : {1}",gp.Key, gp.count)).ToArray," ," & chr(10).ToString),
			Join(Dt_Project_Selector_info.AsEnumerable.GroupBy(Function(row) row("ActivityName").ToString).Select(Function(gp) String.Format("{0} : {1}",gp.Key, gp.count)).ToArray," ," & chr(10).ToString),
			StrArr_Proj_info(3)
		}
		Dt_Project_Activity_info.Rows.Add(StrArr_NewRow)
		'ProjName | StudioVersion | UIAutomaion.Ver | Total_Xaml_Count | Total_UI_Count | Web_Ui_Count | ElementExist_Count | GruopBy_FileName | GruopBy_Activites | Dependencies
	
	Catch ex As Exception
		System.IO.File.WriteAllText("Error_"&Str_ProjectName & ".log", ex.Message)
	End Try
		
	' Loop End Settings
	console.WriteLine(String.Format("{0} / {1} | 추출완료 : {2} 초 소요", int_i.ToString("000"), StrArr_ProjDir.Count.ToString("000"), Now.Subtract(Dtm_StartTime ).TotalSeconds.ToString("0.0000")) )
Next	

'Write CSV
Write_CSV(System.IO.Path.Combine(Str_OutputDir, "Project_info"), Convert_DT_to_CSV(Dt_Project_Activity_info) ) 

'Extract out Arguments
out_Dt_Project_Activity_info = Dt_Project_Activity_info
