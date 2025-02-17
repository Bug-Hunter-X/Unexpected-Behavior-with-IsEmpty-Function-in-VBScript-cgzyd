Function f(a, b)
  If IsEmpty(a) Or IsEmpty(b) Then
    WScript.Echo "Error: a or b is empty"
    Exit Function
  End If
  If VarType(a) <> vbDouble And VarType(a) <> vbInteger Then
      WScript.Echo "Error: Invalid data type for a"
      Exit Function
  End If
  If VarType(b) <> vbDouble And VarType(b) <> vbInteger Then
      WScript.Echo "Error: Invalid data type for b"
      Exit Function
  End If
  c = a + b
  WScript.Echo c
End Function

f(1, 2)
f(1, "")
f("", 2)
f("a", 2)