Attribute VB_Name = "Module29"
Sub バックアップ()
    
    ' 変数を作成
    Dim Save_Filename As String
    
    Dim Save_Dir As String
    
    ' ファイル名を指定
    Save_Filename = "修正一覧表_バックアップ__～" & Format(Now, "yyyyMMdd") & ".xls"
    
    ' 保存するファイルを指定
    If Not Len(Sheets("設定").Range("C3").Value) = 0 Then
    
        Save_Dir = Sheets("設定").Range("C3").Value & "\" & Save_Filename
    Else
        
        Save_Dir = "\\tse\APP\調合\加藤\バックアップ\" & Save_Filename
    End If
    
    
    ' 名前を付けて保存する
    Save_File = Application.GetSaveAsFilename(Save_Dir, _
         FileFilter:="Excelファイル,*.xls,すべてのファイル,*.*")
        
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
