' 설명
' 데이터의 날짜를 열이름으로 박아넣은 DT를 다룰 때 쓰는 모듈.
' ColName에는 날짜, row(key)에는 key값이 있을 떄, 해당 data의 좌표를 엑셀 형식으로 반환
' 가령 0행 0열은 "A0"으로 반환하고, 데이터가 맞지 않으면 환경변수 통해 offset 넣을 수 있음.

' 사전 설정
Environment.SetEnvironmentVariable("Fuc_FindPos_DT1_KeyColName","key")
Environment.SetEnvironmentVariable("Fuc_FindPos_DT1_Offest_Col","1") 
Environment.SetEnvironmentVariable("Fuc_FindPos_DT1_Offest_Row","2") 

'인수화 할 부분
'Dim Str_row As String = "engine2"
'Dim Str_col As String = "2021-01-02"
'Dim dt_tmp As DataTable
'Dim Str_DTName As String = "DT1"

Fuc_FindPos = Function(Str_row As String, Str_col As String, dt_tmp As DataTable, Str_DTName As String)
' Function 내부
Dim Str_KeyColName As String  = Environment.GetEnvironmentVariable( String.Format("Fuc_FindPos_{0}_KeyColName",Str_DTName) )
Dim int_Offset_Col As Integer = CInt(Environment.GetEnvironmentVariable( String.Format("Fuc_FindPos_{0}_Offest_Col",Str_DTName)) )
Dim int_Offset_Row As Integer = CInt(Environment.GetEnvironmentVariable( String.Format("Fuc_FindPos_{0}_Offest_Row",Str_DTName)) )

console.WriteLine("Str_KeyColName : "+Str_KeyColName)
console.WriteLine("int_Offset_Col : "+int_Offset_Col.ToString)
console.WriteLine("int_Offset_Row : "+int_Offset_Row.ToString)

' DT에서 추출한 부분
Dim StrArr_Cols As String() = Enumerable.Range(0,dt_tmp.Columns.Count-1).Select(Function(x) dt_tmp.Columns.Item(x).ColumnName).ToArray
If array.IndexOf(StrArr_Cols, Str_KeyColName) = -1
	Throw New Exception(String.Format("입력된 Dt에 Key Column이 없습니다.{0}{1}", vbnewline, Str_KeyColName))
End If 
Dim StrArr_Rows As String() = dt_tmp.AsEnumerable.Select(Function(x) x( Str_KeyColName ).ToString).ToArray


' 계산할 부분
Dim int_Index_Col As Integer = array.IndexOf(StrArr_Cols, Str_col)
Dim int_Index_Row As Integer = array.IndexOf(StrArr_Rows, Str_row)
Dim Str_Result As String = ""

' 예외 처리
If int_Index_Col = -1
	Throw New Exception(String.Format("열을 찾지 못했습니다.{0}{1}",vbNewLine,Str_col))
End If 
If int_Index_Row = -1
	Throw New Exception(String.Format("행을 찾지 못했습니다.{0}{1}",vbNewLine,Str_row))
End If 

' 좌표 계산
Dim Int_ColNum As Integer = int_Index_Col + int_Offset_Col
Dim Str_ColName As String  = ""

Dim Int_Modulo As Integer
While Int_ColNum > 0
	Int_Modulo =(Int_ColNum-1) Mod 26
	Str_ColName = Convert.ToChar(65 + Int_Modulo).ToString() + Str_ColName
	Int_ColNum = CInt( (Int_ColNum - Int_Modulo) / 26 )
End While

Str_Result  = Str_ColName + (int_Index_Row + int_Offset_Row).ToString
Return Str_Result
End Function

'console.WriteLine(Str_Result)
