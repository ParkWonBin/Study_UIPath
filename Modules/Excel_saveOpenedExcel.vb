	Dim Obj_Application As Object
	
	If in_Int_fileFormat=0 Or String.IsNullOrEmpty(in_Int_fileFormat.ToString) 
		in_Int_fileFormat = 51	'파일 포맷 미지정시 .xlsx 파일로 저장
	End If
	
    On Error GoTo excelNotExist
    Obj_Application = GetObject(, "Excel.Application")
	
    On Error GoTo errHandler
    If Obj_Application IsNot Nothing Then
        With CType(Obj_Application, Application)
            .DisplayAlerts = False
            If String.IsNullOrEmpty(in_Str_wbName) Then	
                With .Workbooks(1)	'WB 명 안 넣었을 시, 실행되어진 엑셀 中 첫번째 엑셀을 저장
                    .SaveAs(Filename:=in_Str_SaveAsPath, FileFormat:=in_Int_fileFormat)
                    .Close()
                End With
                .DisplayAlerts = True
                .Quit()
            Else 'WB 명을 넣었을 시
                With .Workbooks(in_Str_wbName)
                    .SaveAs(Filename:=in_Str_SaveAsPath, FileFormat:=in_Int_fileFormat)
                    .Close()
                End With
                .DisplayAlerts = True
                .Quit()
            End If
        End With
        Console.WriteLine("저장 완료 : " + in_Str_SaveAsPath)
        out_Bln_Result = True

        GoTo endSeq
    Else
        GoTo excelNotExist
    End If

errHandler:
    Console.WriteLine("ERROR : " + Err.Number.ToString + " / " + Err.Description)
	out_Bln_Result = False
    GoTo endSeq

excelNotExist:
	err.Clear
    Console.WriteLine("실행되어진 엑셀이 없습니다.")
    out_Bln_Result = False
    GoTo endSeq

endSeq:
	threading.Thread.Sleep(3000)	'엑셀 종료 後 3초 대기
    Exit Sub
