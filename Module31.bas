Attribute VB_Name = "Module31"
Sub �{������()
Attribute �{������.VB_ProcData.VB_Invoke_Func = "a\n14"
    
    Range("B6").Select
    Range("B5") = Range("B5") / 30
    
    
    
End Sub
Sub ��w�肵�č폜()
    
    Dim num As Integer
     
    If VarType(Range("D2").Value) = 5 Then
    
       num = Range("D2").Value
    
        If num > 4 Then
            Dim mb As VbMsgBoxResult
            mb = MsgBox("�{���ɍ폜���܂����H", vbYesNo + vbQuestion)
            
            If mb = vbYes Then
            
                 Rows(num).Delete
                Range("D2").ClearContents
            End If
            
        Else
    
            MsgBox "�@���̗�͍폜�ł��܂���@�@�@"
        End If
        
    Else
        
        MsgBox " ���l����͂��Ă�������   "
    End If
    
    
    
End Sub
