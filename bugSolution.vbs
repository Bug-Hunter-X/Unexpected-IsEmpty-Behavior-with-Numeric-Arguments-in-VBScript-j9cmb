Function f(a,b)
  If a = vbEmpty Then
    a = 0
  End If
  If b = vbEmpty Then
    b = 0
  End If
  f = a + b
End Function

MsgBox f(1,Empty)