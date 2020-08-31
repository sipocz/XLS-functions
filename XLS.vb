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
out = FormatDate("99991130")
End If

CheckaString = out

End Function

Public Function selectMinDate(input_range As Range, str_selector As String) As String
y = input_range.Rows.Count
x = input_range.Columns.Count

minstr = "9999-12-01"

For Each cell In input_range
    
    numstr = CheckaString(cell.Value, str_selector)
    MsgBox (numstr + "--" + minstr)
    If numstr < minstr Then
    minstr = numstr
    
    End If
    selectMinDate = minstr
Next

End Function
