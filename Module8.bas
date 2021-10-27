Attribute VB_Name = "Module8"
Sub データクリア()
Attribute データクリア.VB_ProcData.VB_Invoke_Func = " \n14"
'
' データクリア Macro
'

'
    Range("B3:B23").ClearContents
    Range("b24").Select
    Selection.ClearContents
    Range("B3").Select
    
End Sub
