Dim Arr_process As process()
Arr_process = Process.GetProcesses.Where(Function(x) x.ProcessName.ToUpper = "EXCEL").toArray
If Arr_process.Count >0
	console.writeline(">>> Kill Excel - 종료 후 2초 대기")
	For Each p As process In Arr_process
		p.kill()
	Next
	System.Threading.Thread.Sleep(2000)	'2초 딜레이
Else 
	console.writeline(">>> Kill Excel - 실행 중인 엑셀 없음.")
End If
