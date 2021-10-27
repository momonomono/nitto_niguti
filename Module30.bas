Attribute VB_Name = "Module30"
Sub データ削除()

    Dim rc As VbMsgBoxResult
    rc = MsgBox("本当に削除してもよろしいですか？", vbYesNo + vbQuestion)
    
    If rc = vbYes Then
    
        Sheets("荷口修正一覧表").Select
        Rows(5).Delete
        
        Sheets("入力シート").Select
        Range("B3").Select
        
    End If
    
End Sub
