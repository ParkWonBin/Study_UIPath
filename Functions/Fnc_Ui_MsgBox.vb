Dim Fnc_Ui_MsgBox As System.Func(Of String, String, Boolean) = Function(Str_Massege As String, Str_title As String) As Boolean
  '2022.04.02|wbpark|message와 title을 입력받아 확인/취소 여부를 bool로 입력받습니다.
  Dim Result As System.Windows.Forms.DialogResult = System.Windows.Forms.MessageBox.Show(Str_Massege, Str_title, MessageBoxButtons.YesNo)
  Return If(Result = System.Windows.Forms.DialogResult.Yes,True,False)
End Function