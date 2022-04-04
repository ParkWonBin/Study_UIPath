'WBpark_20220128_App_Bake_Selector_CSV_From_ProjectXaml
Dim in_Str_ProjPath As String = "D:\H_GA_207_휴양소관리_휴양소_객실_확보_예약_처리"
'------------------------------------------------------------------------------------------------------------------------------------------
'Fnc_Get_All_Files
Dim Fnc_Get_All_Files As System.Func(Of String, String()) = Function(str_path As String) As String()
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
	Console.WriteLine(String.Format(" 파일 : {0} 개 확인", StrArr_Files.Count.ToString))	
	Return StrArr_Files
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
'Fnc_GetAttr 
Dim Fnc_GetAttr As System.Func(Of System.Xml.XmlNode, String, String)  = Function(Node As System.Xml.XmlNode, AttrName As String) As String
    If Node.Attributes.count > 0 Then
        Return If( Enumerable.Range(0, Node.Attributes.count).Select(Function(i) Node.Attributes.ItemOf(i).Name.ToString ).ToArray.Contains(AttrName) , Node.Attributes(AttrName).value.ToString , "")
    Else
        Return ""
    End If
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
'Fnc_Extract_Selector_From_Xaml
Dim Fnc_Extract_Selector_From_Xaml As System.Func(Of String, System.Data.DataTable, System.Data.DataTable) = Function( XamlPath As String, Dt_result As System.Data.DataTable) As System.Data.DataTable
    ' dt 확인
    If Dt_result Is Nothing OrElse Dt_result.Columns.count = 0
        Dt_result = New System.Data.DataTable()
        For Each colName As String In "is_Web_UI|DirectoryName|FileName|ActivityName|DisplayName|Refid|Selector|Selector_Recommended|ReMarks".split("|"c)
                Dt_result.Columns.Add(colName, System.Type.GetType("System.String"))
        Next
    End If

    ' Xaml 읽기 시도.
    Dim doc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
    console.WriteLine("Xaml 읽기 시도 : "+vbNewLine+XamlPath)
    doc.Load(XamlPath)

    ' 데이터 추출 시도
    Dim Str_DirName As String = (New System.IO.FileInfo(XamlPath)).DirectoryName.Split("\"c).last
    Dim Str_FileName As String = (New System.IO.FileInfo(XamlPath)).Name

    ' 전체 요소 선택 > Attribute 있는 요소만 선택 > Selector 있는 요소만 선택
    Dim tags As System.Xml.XmlNodeList = doc.GetElementsByTagName("*")
    Dim Arr_Nodes As System.Xml.XmlNode() = Enumerable.range(0, tags.Count).Select(Function(i) tags(i) ).where(Function(x) Not String.IsNullOrWhiteSpace(Fnc_GetAttr(x, "Selector"))).ToArray
    
    ' 데이터 추출 및 DT에 갱신
    If Arr_Nodes.count > 0 
        ' GetValue, SetValue, TypeInto 등 처리 | 2계층 위
        Dim StrArr_Refid As String() = Arr_Nodes.Select(Function(x) Fnc_GetAttr( x.ParentNode.ParentNode, "sap2010:WorkflowViewState.IdRef") ).toArray
		Dim StrArr_ActivityNames As String() = Arr_Nodes.Select(Function(x) x.ParentNode.ParentNode.Name ).toArray
        Dim StrArr_DisplayNames As String() = Arr_Nodes.Select(Function(x) Fnc_GetAttr( x.ParentNode.ParentNode, "DisplayName") ).toArray
        Dim StrArr_Selectors As String() = Arr_Nodes.Select(Function(x) x.Attributes("Selector").value.Replace(chr(10).ToString,"").Replace(chr(13).ToString,"") ).toArray
		Dim StrArr_Selectors_RCMD As String() = StrArr_Selectors.Select(Function(x) x.Replace("'iexplore.exe'","'msedge.exe'") ).toArray
        Dim StrArr_Is_Web_UI As String() = StrArr_Selectors.Select(Function(x) (x.Contains("<html") OrElse x.Contains("<webctrl") OrElse x.Contains("iexplore") OrElse x.Contains("x:Null") ).ToString ).toArray
        Dim StrArr_ReMarks As String() = Enumerable.Repeat(Of String)("Type A", Arr_Nodes.count).ToArray
        
        ' WaitUiElementVanish, Attach 등 처리 | 0계층 위
        StrArr_Refid = StrArr_ActivityNames.Select(Function(x,i) If( Not String.IsNullOrWhiteSpace(StrArr_DisplayNames(i)) AndAlso x.Contains("ui:"),  StrArr_Refid(i) , Fnc_GetAttr( Arr_Nodes(i),"sap2010:WorkflowViewState.IdRef") ) ).toArray
		StrArr_ReMarks = StrArr_ActivityNames.Select(Function(x,i) If( Not String.IsNullOrWhiteSpace(StrArr_DisplayNames(i)) AndAlso x.Contains("ui:"),  StrArr_ReMarks(i) ,String.Format("Type B{0}BrowserType : {1}{0}UiBrowser : {2}"," | ".ToString,Fnc_GetAttr( Arr_Nodes(i), "BrowserType"),Fnc_GetAttr( Arr_Nodes(i), "UiBrowser") )) ).toArray
		StrArr_DisplayNames = StrArr_ActivityNames.Select(Function(x,i) If( Not String.IsNullOrWhiteSpace(StrArr_DisplayNames(i)) AndAlso x.Contains("ui:"),  StrArr_DisplayNames(i) , Fnc_GetAttr( Arr_Nodes(i),"DisplayName") ) ).toArray
        StrArr_ActivityNames = StrArr_ActivityNames.Select(Function(x,i) If( Not String.IsNullOrWhiteSpace(StrArr_DisplayNames(i)) AndAlso x.Contains("ui:") , x , Arr_Nodes(i).Name ) ).toArray
        
        ' Open, Attach, Browse scope 등 처리 | 6계층 위
		StrArr_Refid = StrArr_ActivityNames.Select(Function(x,i) If( x <> "ui:Target",  StrArr_Refid(i) , Fnc_GetAttr( Arr_Nodes(i).ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode, "sap2010:WorkflowViewState.IdRef")  ) ).toArray
		StrArr_ReMarks = StrArr_ActivityNames.Select(Function(x,i) If( x <> "ui:Target",  StrArr_ReMarks(i) , String.Format("Type C{0}BrowserType : {1}{0}UiBrowser : {2}"," | ".ToString,Fnc_GetAttr( Arr_Nodes(i).ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode, "BrowserType"),Fnc_GetAttr( Arr_Nodes(i).ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode, "UiBrowser") )  ) ).toArray
        StrArr_DisplayNames = StrArr_ActivityNames.Select(Function(x,i) If( x <> "ui:Target",  StrArr_DisplayNames(i) , Fnc_GetAttr( Arr_Nodes(i).ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode, "DisplayName")  ) ).toArray
        StrArr_ActivityNames = StrArr_ActivityNames.Select(Function(x,i) If( x <> "ui:Target" , x , Arr_Nodes(i).ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.Name ) ).toArray
		
        For Each i As Integer In Enumerable.Range(0,StrArr_Selectors.Count)
                Dt_result.Rows.Add({StrArr_Is_Web_UI(i),Str_DirName, Str_FileName, StrArr_ActivityNames(i), StrArr_DisplayNames(i),StrArr_Refid(i), StrArr_Selectors(i),StrArr_Selectors_RCMD(i), StrArr_ReMarks(i)})
        Next
    End If
	
    Return Dt_result
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
'Convert_DT_to_CSV
Dim Convert_DT_to_CSV As System.Func(Of System.Data.DataTable, String) = Function(DT_Source As System.Data.DataTable) As String
	Dim colSep As String =","
	Dim rowSep As String = chr(13).ToString+chr(10).ToString
	Dim StrArr_cols As String() = Enumerable.Range(0,DT_Source.Columns.Count).Select(Function(x) DT_Source.Columns.Item(x).ColumnName).ToArray 
	Dim StrArr_rows As String()  = DT_Source.AsEnumerable.Select(Function(row) Join(row.ItemArray.Select(Function(x) x.ToString.Trim).toArray, colSep) ).ToArray
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
'main
Dim Dt_Proj As New System.Data.DataTable
Dim Str_ProjectPath As String = in_Str_ProjPath
Dim Str_ProjectName As String = Str_ProjectPath.Split("\"c).Where(Function(x) Not String.IsNullOrWhiteSpace(x)).Last

For Each xamlPath As String In Fnc_Get_All_Files(Str_ProjectPath)
    Dt_Proj = Fnc_Extract_Selector_From_Xaml(xamlPath,Dt_Proj)
Next
'Write CSV
Write_CSV(Str_ProjectName, Convert_DT_to_CSV(Dt_Proj) ) 
