Dim Fnc_UI_Get_DirPath As System.Func(Of String, String) = Function(str_Desc As String) As String
  '2022.04.02|wbpark|폴더 선택용 UI
  Dim Dlg_FolderBrowser As System.Windows.Forms.FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog() 
  Dlg_FolderBrowser.Description = str_Desc
  Dlg_FolderBrowser.ShowNewFolderButton = True
  Dlg_FolderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
  
  Dim result As System.Windows.Forms.DialogResult = Dlg_FolderBrowser.ShowDialog()
    If result = System.Windows.Forms.DialogResult.OK Then
    Return Dlg_FolderBrowser.SelectedPath
  End If 
  Return ""
End Function