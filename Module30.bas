Attribute VB_Name = "Module30"
Sub �f�[�^�폜()

    Dim rc As VbMsgBoxResult
    rc = MsgBox("�{���ɍ폜���Ă���낵���ł����H", vbYesNo + vbQuestion)
    
    If rc = vbYes Then
    
        Sheets("�׌��C���ꗗ�\").Select
        Rows(5).Delete
        
        Sheets("���̓V�[�g").Select
        Range("B3").Select
        
    End If
    
End Sub
