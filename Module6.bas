Attribute VB_Name = "Module6"
Sub データ移行()
Attribute データ移行.VB_ProcData.VB_Invoke_Func = " \n14"
'
' データ移行 Macro
    
    ' 荷口修正一覧表に列を増やす
    Sheets("荷口修正一覧表").Select
    Rows(5).Insert
    
    ' 入力データをコピー
    Sheets("入力シート").Select
    Range("B50:B72").Copy
    
    ' 入力データを縦に直して貼り付け
    Sheets("荷口修正一覧表").Select
    Range("B5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=True
    
    ' 入力シートに戻る
    Sheets("入力シート").Select
    Range("B3").Select
    Application.CutCopyMode = False
    
End Sub
