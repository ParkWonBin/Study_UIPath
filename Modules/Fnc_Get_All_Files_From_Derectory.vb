Dim Fnc_Get_All_Files As System.Func(Of String, String()) = Function(str_path As String) As String()
  ' 2022.03.30|wbpak|Get Files
  Dim list_str_dir As New System.Collections.Generic.List(Of String) 
  Dim list_str_file As New System.Collections.Generic.List(Of String)

  list_str_dir.Add(str_path)
  While list_str_dir.Count <> 0
    str_path = list_str_dir.Last
    list_str_dir.RemoveAt(list_str_dir.Count -1)
    list_str_dir.AddRange(System.IO.Directory.GetDirectories(str_path))
    list_str_file.AddRange(System.IO.Directory.GetFiles(str_path))
  End While
  Return list_str_file.ToArray
End Function
