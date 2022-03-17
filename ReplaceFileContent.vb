'Module1.vb
Imports System
Imports System.Linq
Imports System.IO.File

Module Module1
   Sub Main()
    '-----------------------------------------------
    ' 폴더 입력받기 UI
    Dim Fnc_Get_DirPath As System.Func(Of String) = Function() As String
      Dim folderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog() 

      folderBrowserDialog1.Description = "Select the directory that you want to use As the default." 
      folderBrowserDialog1.ShowNewFolderButton = False
      folderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer

      Dim result As System.Windows.Forms.DialogResult = folderBrowserDialog1.ShowDialog()
      If result = System.Windows.Forms.DialogResult.OK Then
              Return FolderBrowserDialog1.SelectedPath
      End If 
      Return ""
    End Function
    '-----------------------------------------------    
    ' 파일 입력받기 UI
    Dim Fnc_Get_FilePath As System.Func(Of String) = Function() As String
      Dim SaveFileDialog1 As System.Windows.Forms.openFileDialog = New System.Windows.Forms.openFileDialog() 
      
      SaveFileDialog1.Filter = "XAML files (*.xaml)|*.xaml|All files (*.*)|*.*"
      SaveFileDialog1.FilterIndex = 2
      saveFileDialog1.RestoreDirectory = True

      Dim result As System.Windows.Forms.DialogResult = saveFileDialog1.ShowDialog()
      If result = System.Windows.Forms.DialogResult.OK Then
                Return saveFileDialog1.FileName
      End If 
      Return ""
    End Function
    '-----------------------------------------------
    ' 폴더 내 모든 파일 조회
    Dim Fnc_Get_All_Files As System.Func(Of String, String()) = Function(str_path As String) As String()
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

    '-----------------------------------------------
    ' 파일 내용 Replace
    Dim Fnc_Replace_File As System.Func(of String, String(), String(), String) = Function(str_source As String, StrArr_Before As String(), StrArr_After As String()) As String
      Dim Str_FileContent As String = System.IO.File.ReadAllText(str_source)
      
      For int_i As Integer = 0 To StrArr_Before.length-1
        Str_FileContent=Str_FileContent.replace(StrArr_Before(int_i),StrArr_After(int_i))
      Next 
      
      System.IO.File.WriteAllText(str_source, Str_FileContent, System.Text.Encoding.UTF8)
      console.WriteLine("Replace 완료 : "+str_source)
      return Str_FileContent    
    End Function 
    '-----------------------------------------------

    '-----------------------------------------------
    ' main
    Dim StrArr_XamlFiles As String() = Fnc_Get_All_Files(Fnc_Get_DirPath())
    Dim Before As String() = {"a","b","c"}
    Dim After As String() = {"A","B","C"}
   
    For Each Xaml_File As String in StrArr_XamlFiles.where(function(x) x.contains(".xaml"))
      Fnc_Replace_File(Xaml_File,Before,After)
    Next
    
  End Sub
End Module

' C:\Windows\Microsoft.NET\Framework64\v4.0.30319\vbc.exe "./Module1.vb"
' ./Module1.exe
