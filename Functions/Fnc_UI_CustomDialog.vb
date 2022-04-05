Dim Fnc_UI_CustomDialog As System.Func(Of String,String,String,String) = Function(caption As String, text As String, selStr As String) As String
  Dim prompt As New System.Windows.Forms.Form With {.Width = 280, .Height = 200, .Text = caption}
  Dim textLabel As New System.Windows.Forms.Label With { .Left = 16, .Top = 20, .Width = 240, .Text = text }
  Dim textBox As New System.Windows.Forms.TextBox With { .Left = 16, .Top = 50, .Width = 240, .TabIndex = 0, .TabStop = True }
  Dim selLabel As New System.Windows.Forms.Label With { .Left = 16, .Top = 130, .Width = 88, .Text = selStr }
  Dim cmbx As New System.Windows.Forms.ComboBox With { .Left = 112, .Top = 130, .Width = 144}
  cmbx.Items.Add("Dark Grey")
  cmbx.Items.Add("Orange")
  cmbx.Items.Add("None")
  cmbx.SelectedIndex = 0
  Dim confirmation As New System.Windows.Forms.Button With { .Text = "In Ordnung!", .Left = 16, .Width = 80, .Top = 88, .TabIndex = 1, .TabStop = True }
  AddHandler confirmation.Click, Sub(sender, e) prompt.Close()
  prompt.Controls.Add(textLabel)
  prompt.Controls.Add(textBox)
  prompt.Controls.Add(selLabel)
  prompt.Controls.Add(cmbx)
  prompt.Controls.Add(confirmation)
  prompt.AcceptButton = confirmation
  prompt.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
  prompt.TopMost=True
  prompt.ShowDialog()
  Return String.Format("{0};{1}", textBox.Text, cmbx.SelectedItem.ToString)
End Function

'Dim Str_tmp As String = Fnc_UI_CustomDialog("caption","text","selStr")
'console.WriteLine(Str_tmp)
              
' New With  :
' https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/objects-and-classes/object-initializers-named-and-anonymous-types
' code : 
' https://stackoverflow.com/questions/5427020/prompt-dialog-in-windows-forms
' https://docs.microsoft.com/ko-kr/dotnet/api/system.windows.forms.combobox.text?view=windowsdesktop-6.0&viewFallbackFrom=dotnet-plat-ext-6.0