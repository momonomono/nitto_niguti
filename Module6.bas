Attribute VB_Name = "Module6"
Sub �f�[�^�ڍs()
Attribute �f�[�^�ڍs.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �f�[�^�ڍs Macro
    
    ' �׌��C���ꗗ�\�ɗ�𑝂₷
    Sheets("�׌��C���ꗗ�\").Select
    Rows(5).Insert
    
    ' ���̓f�[�^���R�s�[
    Sheets("���̓V�[�g").Select
    Range("B50:B72").Copy
    
    ' ���̓f�[�^���c�ɒ����ē\��t��
    Sheets("�׌��C���ꗗ�\").Select
    Range("B5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=True
    
    ' ���̓V�[�g�ɖ߂�
    Sheets("���̓V�[�g").Select
    Range("B3").Select
    Application.CutCopyMode = False
    
End Sub
