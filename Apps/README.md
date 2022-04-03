#### UiPath Invoke Code 주의사항
Invoke Code 안에서는 Sub를 새로 만들 수 없어, 함수를 변수 안에 저장해서 호출하는 방식으로 코드블록을 정리합니다.    
Sub 하나에 붙여넣으면 사용할 수 있는 형태로 코드를 정리하여 올립니다.

dotNet으로 파일을 만들 경우 아래 서식에 붙여넣고 RunVb.Bat을 실행시키면 됩니다.  
```vb

'RunVB.vb
Imports System
Module RunVB
    Public Sub Main()
    '-----------------------------------------------
    'Paste Code Block Here
    '-----------------------------------------------
    end Sub
end Module 
```


#### 자주 사용하는 라이브러리
```vb
System.IO
System.Linq
System.Threading.Thread.Sleep()
System.Diagnostics.Process
' Foreground
Microsoft.VisualBasic.Interaction
System.Windows.Forms
```