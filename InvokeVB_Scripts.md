### Get_AllFiels
파일탐색 
```vb
' Dim in_str_path As String
' Dim out_ArrStr_AllFiles As String[]

Dim list_str_dir  As New List(Of String) 
Dim list_str_file  As New List(Of String)
Dim str_path As String = in_str_path

list_str_dir.Add(str_path)
While list_str_dir.Count <> 0
	str_path = list_str_dir(0)
	list_str_dir.RemoveAt(0)
	list_str_dir.AddRange(System.IO.Directory.GetDirectories(str_path).Where(Function(x) Not x.Contains(".screenshots")).ToArray)
	list_str_file.AddRange(System.IO.Directory.GetFiles(str_path))
    Console.WriteLine(String.Format("남은 폴더 : {0}, 현재 파일 : {1}",list_str_dir.Count.ToString, list_str_file.Count.ToString))
End While

out_ArrStr_AllFiles = list_str_file.ToArray
' join(StrArr_AllFiles.Where(Function(x) x.Contains(".xaml")).ToArray,vbNewLine)

```

### Find UI Activities From Xaml
```vb
'Dim in_str_xmlPath As String
'Dim out_Dt_result As DataTable
Dim doc As System.Xml.XmlDocument
Dim tags As System.Xml.XmlNodeList
Dim Arr_Nodes As System.Xml.XmlNode()

'초기 설정
doc = New XmlDocument()
doc.Load(in_str_xmlPath)

' 전체 요소 선택 > Attribute 있는 요소만 선택 > Selector 있는 요소만 선택
tags  = doc.GetElementsByTagName("*")
Arr_Nodes = Enumerable.range(0, tags.Count).Select(Function(i) tags(i) ).where(Function(x) x.Attributes.count>0).toArray
Arr_Nodes = Arr_Nodes.where(Function(x) enumerable.range(0, x.Attributes.count).Select(Function(i) x.Attributes.ItemOf(i).Name.ToString ).ToArray.Contains("Selector")).ToArray

Dim Str_DirName As String
Dim Str_FileName As String
Dim StrArr_ActivityNames As String()
Dim StrArr_DisplayNames As String()
Dim StrArr_Selectors As String()
Dim StrArr_Is_Web_UI As String()

Str_DirName = (new System.IO.FileInfo(in_str_xmlPath)).DirectoryName.Split("\"c).last
Str_FileName = New system.io.FileInfo(in_str_xmlPath).Name
StrArr_ActivityNames = Arr_Nodes.Select(Function(x) x.ParentNode.ParentNode.Name ).toArray
StrArr_DisplayNames = Arr_Nodes.Select(Function(x) x.ParentNode.ParentNode.Attributes("DisplayName").value ).toArray
StrArr_Selectors = Arr_Nodes.Select(Function(x) x.Attributes("Selector").value ).toArray
StrArr_Is_Web_UI = StrArr_Selectors.Select(Function(x) (x.Contains("<html") OrElse x.Contains("<webctrl") OrElse x.Contains("iexplore")).ToString ).toArray

'Dim out_Dt_result As DataTable
out_Dt_result = New DataTable()

'열추가
For Each colName As String In "DirectoryName|FileName|ActivityName|DisplayName|Selector|is_Web_UI".split("|"c)
	out_Dt_result.Columns.Add(colName, System.Type.GetType("System.String"))
Next

'OpenBrowser
For Each tag As System.Xml.XmlNode In doc.GetElementsByTagName("ui:OpenBrowser")
	out_Dt_result.Rows.Add({Str_DirName, Str_FileName, tag.Name, tag.Attributes("DisplayName").value, "UiBrowser : " & tag.Attributes("UiBrowser").Value,"TRUE" })
Next

'Selector 존재하는 데이터 추가
For Each i As Integer In Enumerable.Range(0,StrArr_Selectors.Count)
	out_Dt_result.Rows.Add({Str_DirName, Str_FileName, StrArr_ActivityNames(i), StrArr_DisplayNames(i), StrArr_Selectors(i), StrArr_Is_Web_UI(i)})
Next
```
