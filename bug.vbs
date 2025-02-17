Function f(a, b)
  If IsEmpty(a) Then
    WScript.Echo "Error: a is empty"
    Exit Function
  End If
  If IsEmpty(b) Then
    WScript.Echo "Error: b is empty"
    Exit Function
  End If
  c = a + b
  WScript.Echo c
End Function

f(1, 2)