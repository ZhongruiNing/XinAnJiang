Attribute VB_Name = "chhthf"
Function ffq(t As Single, tx() As Single, q() As Single, nn As Integer) As Single
Dim dq As Single, i As Integer
If t <= tx(0) Or t >= tx(nn) Then
    If t <= tx(0) Then
    ffq = q(0)
    Else
    ffq = q(nn)
    End If
Else
  For i = 0 To nn - 1
    If t >= tx(i) And t <= tx(i + 1) Then
        dq = Abs(q(i + 1) - q(i))
         If dq < 1# Then
         ffq = q(i)
         Else
         ffq = q(i) + (t - tx(i)) / (tx(i + 1) - tx(i)) * (q(i + 1) - q(i))
         End If
    Exit For
    End If
  Next i
End If
End Function


