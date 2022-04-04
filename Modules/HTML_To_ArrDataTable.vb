Dim Fnc_Get_Html_Tags As System.Func(Of String,String,String, String()) = Function(Source As String, TagName As String, contents As String) As String()
	Return System.Text.RegularExpressions.Regex.Split(Source,String.format("<{0}[^>]*>|</{0}[^>]*>",TagName),System.Text.RegularExpressions.RegexOptions.IgnoreCase).Where(Function(x) Not String.IsNullOrWhiteSpace(x) AndAlso System.Text.RegularExpressions.Regex.Matches(x,contents).Count>0 ).ToArray
End Function

Dim Fnc_Extract_Html_Tabels As System.Func(Of String, System.Data.DataTable()) = Function(HTML As String) As System.Data.DataTable()
    Dim Result As New List(Of Data.DataTable)
    For Each table As String In Fnc_GetTags(HTML,"table","</tr>")
        Dim data As String()() = Fnc_GetTags(table,"tr","[</th>|</td>]").Select(Function(tr) Fnc_GetTags(tr,"[th|td]","") ).ToArray

        Dim dt_tmp As New DataTable
		'Add Columns
        For Each colName As String In data.first 
            dt_tmp.Columns.Add(colName, System.Type.GetType("System.String")) 
        Next
		 'Add Data
        For Each NewRow As String() In data.skip(1)
            dt_tmp.Rows.Add(NewRow)
        Next
		
        Result.Add(dt_tmp)
    Next
    Return Result.ToArray
End Function

out_Fuc_GetTags = Fnc_GetTags
out_Fnc_Extract_Tabels = Fnc_Extract_Tabels
