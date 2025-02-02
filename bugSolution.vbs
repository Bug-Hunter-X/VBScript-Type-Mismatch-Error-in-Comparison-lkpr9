Function f(a, b)
  ' Convert both variables to numbers before comparison
  Dim aNum, bNum
  aNum = CDbl(a)
  bNum = CDbl(b)
  If aNum > bNum Then
    MsgBox "a is greater than b"
  ElseIf aNum < bNum Then
    MsgBox "a is less than b"
  Else
    MsgBox "a is equal to b"
  End If
End Function

' This will now work correctly
f(1, "2")
f(2, 1)
f("2", 2) 