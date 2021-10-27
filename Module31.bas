Attribute VB_Name = "Module31"
Sub 本数入力()
Attribute 本数入力.VB_ProcData.VB_Invoke_Func = "a\n14"
    
    Range("B6").Select
    Range("B5") = Range("B5") / 30
    
    
    
End Sub
Sub 列指定して削除()
    
    Dim num As Integer
     
    If VarType(Range("D2").Value) = 5 Then
    
       num = Range("D2").Value
    
        If num > 4 Then
            Dim mb As VbMsgBoxResult
            mb = MsgBox("本当に削除しますか？", vbYesNo + vbQuestion)
            
            If mb = vbYes Then
            
                 Rows(num).Delete
                Range("D2").ClearContents
            End If
            
        Else
    
            MsgBox "　その列は削除できません　　　"
        End If
        
    Else
        
        MsgBox " 数値を入力してください   "
    End If
    
    
    
End Sub
