Dim str_tmp As String=""
Dim folderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog() 

folderBrowserDialog1.Description = "Select the directory that you want to use As the default." 
folderBrowserDialog1.ShowNewFolderButton = False
folderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer

Dim result As System.Windows.Forms.DialogResult = folderBrowserDialog1.ShowDialog()
If result = System.Windows.Forms.DialogResult.OK Then
        str_tmp = FolderBrowserDialog1.SelectedPath
End If 
console.WriteLine(str_tmp)
