Attribute VB_Name = "Module29"
Sub �o�b�N�A�b�v()
    
    ' �ϐ����쐬
    Dim Save_Filename As String
    
    Dim Save_Dir As String
    
    ' �t�@�C�������w��
    Save_Filename = "�C���ꗗ�\_�o�b�N�A�b�v__�`" & Format(Now, "yyyyMMdd") & ".xls"
    
    ' �ۑ�����t�@�C�����w��
    If Not Len(Sheets("�ݒ�").Range("C3").Value) = 0 Then
    
        Save_Dir = Sheets("�ݒ�").Range("C3").Value & "\" & Save_Filename
    Else
        
        Save_Dir = "\\tse\APP\����\����\�o�b�N�A�b�v\" & Save_Filename
    End If
    
    
    ' ���O��t���ĕۑ�����
    Save_File = Application.GetSaveAsFilename(Save_Dir, _
         FileFilter:="Excel�t�@�C��,*.xls,���ׂẴt�@�C��,*.*")
        
     ActiveWorkbook.SaveAs Filename:=Save_File, _
                           FileFormat:=xlOpenXMLWorkbookMacroEnabled, _
                           ReadOnlyRecommended:=False, _
                           CreateBackup:=False

                           

    
         
End Sub
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    
    
End Sub
