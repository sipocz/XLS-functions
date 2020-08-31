Public Function FindNum(str_input As String) As String

s = ""
For i = 1 To Len(str_input)
    ch = Mid(str_input, i, 1)
    If IsNumeric(ch) Then
        s = s + ch
    End If
        
Next i

FindNum = s

End Function

Public Function FormatDate(str_input As String) As String
Dim o As String
If Len(str_input) > 6 Then
ys = Mid(str_input, 1, 4)
ms = Mid(str_input, 5, 2)
ds = Mid(str_input, 7, 2)
Else
ys = Mid(str_input, 1, 2)
ms = Mid(str_input, 3, 2)
ds = Mid(str_input, 5, 2)
End If

o = ys + "-" + ms + "-" + ds

FormatDate = o
End Function

Public Function CheckaString(str_input As String, str_selector As String) As String
Dim out As String
If InStr(str_input, str_selector) Then
out = FormatDate(FindNum(str_input))
Else
out = FormatDate("99991231")
End If

CheckaString = out

End Function

Public Function GetaRange(r as range) As String 
for 

End Function