Attribute VB_Name = "Module20"
Public Function myf(a) As String
Dim i As Long
For i = 1 To Len(a)
If Mid(a, i, 1) Like "[/a-zA-Z]" Then
 myf = myf & Mid(a, i, 1)
End If
Next i
End Function

Public Function myg(a) As String
Dim i As Long
For i = 1 To Len(a)
If Mid(a, i, 1) Like "[0-9]" Then
 myg = myg & Mid(a, i, 1)
End If
Next i


End Function
