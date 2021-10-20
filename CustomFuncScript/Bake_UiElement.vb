'Indicate On Screen =>  ui_tmp As UiPath.Core.UiElement 
Dim dic_attar As dictionary(Of String, String)
Dim Str_result As String  = ""

dic_attar = ui_tmp.GetNodeAttributes(False)
For Each key As String In dic_attar.keys()
	'console.writeline(key+" : "+dic_attar(key))
	Str_result = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,key,dic_attar(key))
Next
Str_result  = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,"Selector",ui_tmp.Selector.Text)
Str_result  = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,"SelectorStrategy",ui_tmp.SelectorStrategy.ToString)
Str_result  = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,"ParentSelector",ui_tmp.Parent.Selector.Text)
Str_result  = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,"TopParent",ui_tmp.TopParent().Selector.Text)

Dim Str_fileName As String = Environment.CurrentDirectory+"\UiAttar.md" 
If file.Exists(Str_fileName) Then 
	file.delete(Str_fileName )
End If 
file.WriteAllText(Str_fileName ,Str_result)
