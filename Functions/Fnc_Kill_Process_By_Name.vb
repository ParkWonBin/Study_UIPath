Dim Fnc_Kill_Process_By_Name As System.Func(Of String, String) = Function(ProcessName As String) As String
  '2022.04.04|wbpark|DRM 있는 엑셀 종료를 위해 제작
  Dim Arr_process As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcesses.Where(Function(x) x.ProcessName.ToUpper = ProcessName.ToUpper).toArray
    While Arr_process.Count>0
        For Each p As System.Diagnostics.Process In Arr_process
            p.kill()
        Next
        System.Threading.Thread.Sleep(1000)	'1초 딜레이
        Arr_process=System.Diagnostics.Process.GetProcesses.Where(Function(x) x.ProcessName.ToUpper = ProcessName.ToUpper).toArray
    End While 
    Return ProcessName+" 종료"
End Function