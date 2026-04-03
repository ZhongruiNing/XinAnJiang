Attribute VB_Name = "variety2"



Sub inputparad(para1() As Single, para2() As Single, para3() As String, Na As Integer, ia As Integer)
Dim i As Integer, j As Integer

bname = "cscd" + dyly + "1"
sql1 = "select * from " + bname + "  ORDER BY [序号]"
rd1.CursorLocation = adUseClient
rd1.Open sql1, cn

For i = 1 To rd1.RecordCount - 1
   para1(i) = Val(rd1(3))
   rd1.MoveNext
Next i
   Zhanevap = rd1(3)
   rd1.MoveNext
rd1.Close

Na = para1(18): ia = para1(19): Ma = para1(21)
ReDim para2(2, Na + ia), para3(Na + ia)

bname = "cscd" + dyly + "2"
sql1 = "select * from " + bname + "  ORDER BY [序号]"
rd1.CursorLocation = adUseClient
rd1.Open sql1, cn

If Not rd1.BOF Then
rd1.MoveFirst
End If
Do While Not rd1.EOF
i = rd1(0)
para3(i) = rd1(1)
para2(1, i) = rd1(2)
para2(2, i) = rd1(3)
rd1.MoveNext
Loop
rd1.Close

If Ma > 1 Then
   ReDim para4(Ma, 3)
   bname = "cscd" + dyly + "3"
   rd1.CursorLocation = adUseClient
   rd1.Open bname, cn
   
   For i = 1 To Ma
      For j = 1 To 3
         para4(i, j) = rd1(j)
      Next j
      rd1.MoveNext
   Next i
   rd1.Close
Else
End If

End Sub
Sub inputparah(para1() As Single, para2() As Single, para3() As String, Na As Integer, ia As Integer)
Dim i As Integer
bname = "csch" + dyly + "1"
rd1.CursorLocation = adUseClient
rd1.Open bname, cn


If Not rd1.BOF Then
rd1.MoveFirst
End If
Do While Not rd1.EOF
i = rd1(0)
para1(i) = Val(rd1(3))
rd1.MoveNext
Loop
rd1.Close
Na = para1(18): ia = para1(19)
ReDim para2(2, Na + ia), para3(Na + ia)

bname = "csch" + dyly + "2"
rd1.Open bname, cn

If Not rd1.BOF Then
rd1.MoveFirst
End If
Do While Not rd1.EOF
i = rd1(0)
para3(i) = rd1(1)
para2(1, i) = rd1(2)
para2(2, i) = rd1(3)
rd1.MoveNext
Loop
rd1.Close
End Sub
Sub eevap(evap() As Single)

Dim i As Integer, j As Integer, k As Integer, it As Long, nn As Integer

Dim ev As Single, revp() As Single
Dim mn As Integer
nn = glday
ReDim revp(nn)

Open pathc + "\evap.dat" For Input As #1
Input #1, ev
Close #1
For i = 0 To nn
revp(i) = ev
Next i
bname = "devapo" + dyly
it1 = sdsj(0)
it2 = sdsj(glday)
sql1 = "select * from " + bname + " where dt between  " + CStr(it1) + " and " + CStr(it2) + "  Order by dt"

b.CursorLocation = adUseClient
b.Open sql1, cn

 If Not b.BOF Then
          For j = 1 To nn
             If b(1) >= 0.0001 Then
                revp(j) = b(1)
             Else
               revp(j) = 0
             End If
             b.MoveNext
          Next j
      End If

b.Close
mn = Int(24 / gltt)

For i = 1 To nn
  j = (i - 1) * mn
   For k = 1 To mn
      evap(j + k) = revp(i) / mn
   Next k
Next i

End Sub


Sub eevapD(evap() As Single)
Dim i As Integer, j As Integer, k As Integer, it As Long, nn As Integer

Dim ev As Single, revp() As Single
Dim mn As Integer
nn = LongTimeD
ReDim revp(nn)

Open pathc + "\evap.dat" For Input As #1
Input #1, ev
Close #1
For i = 0 To nn
revp(i) = ev
Next i
bname = "devapo" + dyly
it1 = sdsj(1)
it2 = sdsj(glday)

sql1 = "select " + CStr(Zhanevap) + " from " + bname + " where dt between  " + CStr(it1) + " and " + CStr(it2) + "  order by DT "

b.CursorLocation = adUseClient
b.Open sql1, cn

 If Not b.BOF Then
          For j = 1 To glday
             If b(0) >= 0.0001 Then
                revp(j) = b(0)
             Else
                revp(j) = 0
             End If
             b.MoveNext
          Next j
      End If

b.Close
For i = 1 To nn
   evap(i) = revp(i)
Next i

End Sub

Sub stateh(Na As Integer, state() As Single)
 
Dim i As Integer, it As Long, itm As Long, itt As Long, j As Integer, ii As Integer
Dim yy As Integer, mm As Integer, dd As Integer, na1 As Integer
Dim iy1 As Integer, im1 As Integer, id1 As Integer, ih1 As Integer
Dim sumpara() As Single, Rcount As Integer, wt() As Single, sumpa() As Single
Dim msg, Style, Response
Dim prartest() As Single

ReDim paratest(1 To 23)

it = StartTime
Call ymdh(it, iy1, im1, id1, ih1)

Call tymd(iy1, im1, id1, -1, yy, mm, dd)
Call yrsfd(yy, mm, dd, itm)
Call yrsfd(iy1, im1, id1, itt)


bname = "cscd" + dyly + "1"
sql1 = "select * from " + bname + "  ORDER BY [序号]"
rd1.CursorLocation = adUseClient
rd1.Open sql1, cn

For i = 1 To rd1.RecordCount - 1
   paratest(i) = Val(rd1(3))
   rd1.MoveNext
Next i
   Zhanevap = rd1(3)
   rd1.MoveNext

na1 = paratest(18)
rd1.Close
''''
ReDim wt(Na, 7)

bname = "dast" + dyly
sql1 = "select * from " + bname + " where dt = " + CStr(itm)
rd1.CursorLocation = adUseClient
rd1.Open sql1, cn
'****************************************************************

If rd1.EOF Then
  rd1.Close
  msg = dylyc + "无日模初始值,请选择采用初始化(是）或者进行日模计算（否）！"
  glztz = 0
  Style = vbYesNo
  Response = MsgBox(msg, Style)
  If Response = vbYes Then
    bname = "dast" + dyly + "0"
    rd2.CursorLocation = adUseClient
    rd2.Open bname, cn
  
    If Not rd2.EOF = True Then
        For i = 1 To 6
            state(0, i) = rd2(i - 1)
        Next i
        rd2.Close
    Else
       MsgBox CStr(iy1) + dylyc + "无初始化值，请先建立初始值文件！"
       Exit Sub
    End If
  Else
   Call chyubasW
   Exit Sub
 End If
Else
   glztz = 1
 


    ReDim sumpara(1 To 7), sumpa(1 To 7)
    If na1 > Na Then
        ReDim state(na1, 7)
    End If

'na1日资料雨量站个数;na次资料雨量站个数两者不等时进行处理
    rd1.MoveFirst
    'rd1.MoveNext
   For j = 1 To na1 + 1
      ii = rd1(2)
      For i = 1 To 7
        If rd1(i + 4) > 0.001 Then
          state(ii, i) = rd1(i + 4)
          sumpara(i) = sumpara(i) + state(ii, i)
        Else
          state(ii, i) = 0
          sumpara(i) = sumpara(i) + state(ii, i)
        End If
      Next i
    rd1.MoveNext
  Next j
    rd1.Close
End If


    If Na <> na1 Then
         For ii = 1 To Na
            For i = 1 To 7
                state(ii, i) = sumpara(i) / na1
            Next i
         Next ii
    End If
    

'----------------------------------------------------------
it = StartTime
Call ymdh(it, iy1, im1, id1, ih1)
If ih1 > 8 Then
Dim wtt As Single

bname = "dast" + dyly
sql1 = "select * from " + bname + " where dt = " + CStr(itt)
rd1.CursorLocation = adUseClient
rd1.Open sql1, cn
rd1.MoveFirst
  For j = 1 To na1 + 1
      ii = rd1(2)
      For i = 1 To 7
        If rd1(i + 4) > 0.001 Then
          wt(ii, i) = rd1(i + 4)
          sumpa(i) = sumpa(i) + wt(ii, i)
        Else
          wt(ii, i) = 0
          sumpa(i) = sumpa(i) + wt(ii, i)
        End If
      Next i
    rd1.MoveNext
  Next j
    rd1.Close
    
    wtt = (ih1 - 8) / 24
      For ii = 1 To Na
            For i = 1 To 7
                state(ii, i) = state(ii, i) + (wt(ii, i) - state(ii, i)) * wtt
                If state(ii, i) = 0 Then state(ii, i) = 0
            Next i
      Next ii
    
     If Na <> na1 Then
         For ii = 1 To Na
            For i = 1 To 7
                state(ii, i) = sumpa(i) / na1
            Next i
         Next ii
    End If

Else
End If

'----------------------------------------------------------
it = StartTime
Call ymdh(it, iy1, im1, id1, ih1)

Call tymd(iy1, im1, id1, -2, yy, mm, dd)
Call yrsfd(yy, mm, dd, itt)

If ih1 < 8 Then


bname = "dast" + dyly
sql1 = "select * from " + bname + " where dt = " + CStr(itt)
rd1.CursorLocation = adUseClient
rd1.Open sql1, cn
rd1.MoveFirst
  For j = 1 To na1 + 1
      ii = rd1(2)
      For i = 1 To 7
        If rd1(i + 4) > 0.001 Then
          wt(ii, i) = rd1(i + 4)
          sumpa(i) = sumpa(i) + wt(ii, i)
        Else
          wt(ii, i) = 0
          sumpa(i) = sumpa(i) + wt(ii, i)
        End If
      Next i
    rd1.MoveNext
  Next j
    rd1.Close
    
    wtt = (8 - ih1) / 24
      For ii = 1 To Na
            For i = 1 To 7
                state(ii, i) = state(ii, i) + (wt(ii, i) - state(ii, i)) * wtt
                If state(ii, i) = 0 Then state(ii, i) = 0
            Next i
    Next ii
    
     If Na <> na1 Then
         For ii = 1 To Na
            For i = 1 To 7
                state(ii, i) = sumpa(i) / na1
            Next i
         Next ii
    End If

Else
End If

End Sub

Sub statehD(Na As Integer, state() As Single)
 
Dim i As Integer, it As Long, itm As Long, j As Integer, ii As Integer
Dim yy As Integer, mm As Integer, dd As Integer
Dim iy1 As Integer, im1 As Integer, id1 As Integer
Dim msg, Style, Response

it = StartTimeD
Call ymd(it, iy1, im1, id1)

Call tymd(iy1, im1, id1, -1, yy, mm, dd)
Call yrsfd(yy, mm, dd, itm)

bname = "dast" + dyly
sql1 = "select * from  " + bname + " where dt =" + CStr(itm)
rd1.CursorLocation = adUseClient
rd1.Open sql1, cn

If rd1.EOF Then
  rd1.Close
'  msg = dylyc + "无日模初始值,请选择采用初始化(是）或者停止计算（否）！"
  glztz = 0
'  Style = vbYesNo
'  Response = MsgBox(msg, Style)
' If Response = vbYes Then
    bname = "dast" + dyly + "0"
    rd2.CursorLocation = adUseClient
    rd2.Open bname, cn
   If Not rd2.EOF = True Then
      For i = 1 To 6
          state(0, i) = rd2(i - 1)
      Next i
      rd2.Close
    Else
       MsgBox CStr(iy1) + dylyc + "无初始化值，请先建立初始值文件！"
       rd2.Close
       Exit Sub
    End If
'  Else
'    Exit Sub
'  End If
Else
  glztz = 1
  For j = 0 To Na
    ii = rd1(2)
    For i = 1 To 7
      If rd1(i + 4) > 0.001 Then
        state(ii, i) = rd1(i + 4)
      Else
        state(ii, i) = 0
      End If
    Next i
    rd1.MoveNext
   Next j
rd1.Close
End If

End Sub




Sub inputrain(Na As Integer, zdylp() As Single, para3() As String)

'  na.....the number of sub_basins
' zdylp..the data of rainfall for every subbasin

Dim i As Integer, j As Integer, jj As Integer, it As Long
Dim it1 As Long, it2 As Long

Dim inname() As String
ReDim inname(Na)
 
 bname = "rain_period" + dyly
 it1 = glchsdsj(1)
 it2 = glchsdsj(glnn1)
 
For i = 1 To Na
    inname(i) = para3(i)

    sql1 = "select " + CStr(inname(i)) + " from " + bname + " where dt between  " + CStr(it1) + " and " + CStr(it2) + "  order by DT "
    b.CursorLocation = adUseClient
    b.Open sql1, cn
    If b.BOF Or b.EOF Then
       MsgBox dylyc + "日雨量文件不存在，请建立该文件！"
       b.Close
       Exit Sub
    Else
       For j = 1 To glnn1
           If b(0) > 0.0001 Then
              zdylp(j, i) = b(0)
           Else
              zdylp(j, i) = 0
           End If
           b.MoveNext
      Next j
  End If
  b.Close
Next i

For j = glnn1 + 1 To glnn
   For i = 1 To Na
     zdylp(j, i) = 0
   Next i
Next j
End Sub
Sub inputrainD(Na As Integer, zdylp() As Single, para3() As String)

'  na.....the number of sub_basins
' zdylp..the data of rainfall for every subbasin

Dim i As Integer, j As Integer, jj As Integer, it As Long
Dim it1 As Long, it2 As Long

Dim inname() As String
ReDim inname(Na)
 
 bname = "rain_day" + dyly
 it1 = sdsj(1)
 it2 = sdsj(glnn1)
 
For i = 1 To Na
    inname(i) = para3(i)

    sql1 = "select " + CStr(inname(i)) + " from " + bname + " where dt between  " + CStr(it1) + " and " + CStr(it2) + "  order by DT "
    b.CursorLocation = adUseClient
    b.Open sql1, cn
    If b.BOF Or b.EOF Then
       MsgBox dylyc + "日雨量文件不存在，请建立该文件！"
       b.Close
       Exit Sub
    Else
       For j = 1 To glnn1
           If b(0) > 0.0001 Then
              zdylp(j, i) = b(0)
           Else
              zdylp(j, i) = 0
           End If
           b.MoveNext
      Next j
  End If
  b.Close
Next i

For j = glnn1 + 1 To glnn
   For i = 1 To Na
     zdylp(j, i) = 0
   Next i
Next j
End Sub


Sub inputq(qobs() As Single)
' ia.....the number of inflow
' zdylq..the data of inflow from reservoir or lake
' this sub is only for bt xx hc jj qj
Dim i As Integer, j As Integer, it As Long
Dim it1 As Long, it2 As Long

 bname = "qdata" + dyly
it1 = glchsdsj(1)
it2 = glchsdsj(glnn1)

sql1 = "select dt," + CStr(dylyc) + " from " + bname + " where dt between  " + CStr(it1) + " and " + CStr(it2) + "  order by DT "
 b.CursorLocation = adUseClient
 b.Open sql1, cn

If b.BOF Or b.EOF Then
   MsgBox dylyc + "次洪实测流量文件不存在，请建立该文件！"
   Exit Sub
Else
End If
For j = 1 To glnn1
     qobs(j) = b(1)
    b.MoveNext
Next j
b.Close

End Sub
Sub inputqD(qobs() As Single)

' ia.....the number of inflow
' zdylq..the data of inflow from reservoir or lake
' this sub is only for bt xx hc jj qj
Dim i As Integer, j As Integer, it As Long
Dim it1 As Long, it2 As Long


If dyly <> "lm" Then

bname = "qdata_day" + dyly
 
it1 = sdsj(1)
it2 = sdsj(glnn)

sql1 = "select " + CStr(dylyc) + " from " + bname + " where dt between  " + CStr(it1) + " and " + CStr(it2) + "  order by DT "
 b.CursorLocation = adUseClient
 b.Open sql1, cn

If b.BOF Or b.EOF Then
   MsgBox dylyc + "次洪实测流量文件不存在，请建立该文件！"
   Exit Sub
Else
End If
j = 1
Do While Not b.EOF
     If b(0) > 0.0001 Then
      qobs(j) = b(0)
    Else
      qobs(1) = 0
    End If
    b.MoveNext
    j = j + 1
Loop
b.Close

For i = 1 To m
   If qobs(i) < 0 Then
         qobs(i) = 0
   Else
   End If

Next i

Else
   For i = 1 To m
         qobs(i) = 0
   Next i
End If

Qjliu = qobs(1)
End Sub






