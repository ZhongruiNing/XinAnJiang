Attribute VB_Name = "readq"
Sub readqq()
Dim i As Integer, j As Integer, it As Long
Dim it1 As Long, it2 As Long, sql1 As String
Dim q() As Single
ReDim q(2, lontime)

it1 = StartTime
it2 = EndTime
bname = "qhredata"
sql1 = "select * from " + bname + " where dt between  " + Str(it1) + " and " + Str(it2) + "  order by DT  "
rst.CursorLocation = adUseClient
rst.Open sql1, cn

If Not rst.BOF Then
rst.MoveFirst
End If
Do While Not rst.EOF
For j = 1 To lontime
      q(1, j) = rst(1)
      q(2, j) = rst(2)
rst.MoveNext
Next j
Loop
rst.Close

ReDim glpwhf(LongTime), glqcalhf(LongTime), glqobshf(LongTime), glqadjhf(LongTime)
For i = 1 To LongTime
  glpwhf(i) = 0#
  glqcalhf(i) = q(1, i)
  glqobshf(i) = q(2, i)
  glqadjhf(i) = 0#
Next i

End Sub
Sub readfore()
Dim i As Integer, j As Integer, it As Long
Dim it1 As Long, it2 As Long

ReDim glpwhf(LongTime), glqcalhf(LongTime), glqobshf(LongTime), glqadjhf(LongTime), sdsj(LongTime)

If Showw = "MX" Then
it1 = StartTime
it2 = EndTime
bname = "ybresu" + dyly
sql1 = "select * from " + bname + " where dt between  " + CStr(it1) + " and " + CStr(it2) + "  order by DT  "
rst.CursorLocation = adUseClient
rst.Open sql1, cn

If Not rst.BOF Then
rst.MoveFirst
For j = 1 To LongTime
      sdsj(j) = rst(2)
      glpwhf(j) = rst(3)
      glqcalhf(j) = rst(9)
      glqobshf(j) = rst(8)
      glqadjhf(j) = rst(6)
     
      rst.MoveNext
Next j
End If
rst.Close

Else

it1 = StartTime
it2 = EndTime
bname = "jyresu" + dyly
sql1 = "select * from " + bname + " where dt between  " + CStr(it1) + " and " + CStr(it2) + "  order by DT  "
rst.CursorLocation = adUseClient
rst.Open sql1, cn

If Not rst.BOF Then
rst.MoveFirst
For j = 1 To LongTime
      sdsj(j) = rst(1)
      glpwhf(j) = rst(4)
      glqcalhf(j) = rst(7)
      glqobshf(j) = rst(6)
      glqadjhf(j) = rst(8)
      rst.MoveNext
Next j
End If
rst.Close

End If

End Sub

Sub readfored()
Dim i As Integer, j As Integer, it As Long
Dim it1 As Long, it2 As Long

ReDim glpwhf(LongTimeD), glqcalhf(LongTimeD), glqobshf(LongTimeD), glqadjhf(LongTimeD)

it1 = StartTimeD
it2 = EndTimeD
bname = "ybresuD" + dyly
sql1 = "select * from " + bname + " where Äê·Ý=" + CStr(NoYear) + "and dt between  " + CStr(it1) + " and " + CStr(it2) + "  order by DT  "
rst.CursorLocation = adUseClient
rst.Open sql1, cn
If Not rst.BOF Then
rst.MoveFirst
For j = 1 To LongTimeD
      sdsj(j) = rst(1)
      glpwhf(j) = rst(2)
      glqcalhf(j) = rst(8)
      glqobshf(j) = rst(7)
      glqadjhf(j) = rst(7)
      rst.MoveNext
Next j
End If
rst.Close
End Sub


