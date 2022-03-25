'Module1.vb
Imports System

Module RunVB1
    Public Sub Main()
    '-----------------------------------------------
    dim str_tmp as string
    str_tmp = "123"
    Microsoft.VisualBasic.Interaction.MsgBox("MsgBox(123) 입력 : " +str_tmp)
    Console.WriteLine("Console.WriteLine(""123"") 입력 : "+str_tmp)
    '-----------------------------------------------
    end Sub
end Module 
