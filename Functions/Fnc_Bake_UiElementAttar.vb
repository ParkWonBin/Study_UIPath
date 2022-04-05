Dim Fnc_Bake_UiElementAttar As System.Func(Of UiPath.Core.UiElement, String) = Function(ui_tmp As UiPath.Core.UiElement) As String
  '2022.04.04|wbpark|Uipath Activity와 함께 써야하는 함수. Invoke Code 내에서 Uipath Core 함수 호출 방법 불명.
  Dim Str_result As String  = ""
  Dim Str_fileName As String = Environment.CurrentDirectory+"\UiElementAttar.log"
  Dim dic_attar As dictionary(Of String, String) = ui_tmp.GetNodeAttributes(False)
  For Each key As String In dic_attar.keys()
    Str_result = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,key,dic_attar(key))
  Next
  Str_result  = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,"Selector",ui_tmp.Selector.Text)
  Str_result  = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,"SelectorStrategy",ui_tmp.SelectorStrategy.ToString)
  Str_result  = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,"ParentSelector",ui_tmp.Parent.Selector.Text)
  Str_result  = String.format("{0}{2} : {3}{1}", Str_result, vbnewline,"TopParent",ui_tmp.TopParent().Selector.Text)
  System.IO.File.WriteAllText(Str_fileName ,Str_result)
  Return Str_result
End Function