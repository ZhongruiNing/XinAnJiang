Attribute VB_Name = "fenpjun"
Option Explicit
Public ResultCn As New ADODB.Connection
Public ResultRs As New ADODB.Recordset
Sub findyear(CountFlood As Integer)

Dim i As Integer, j As Integer

Dim it As Long, im As Integer, id As Integer
Dim yy0 As Integer, mm0 As Integer, dd0 As Integer, hh0 As Integer
Dim yy1 As Integer, mm1 As Integer, dd1 As Integer, hh1 As Integer

bname = "dataflood"

b.CursorLocation = adUseClient
b.Open bname, cn

If Not b.BOF Then
b.MoveFirst
End If
Do While Not b.EOF
i = b(0)
If i = FloodNo(1, CountFlood) Then
      NoYear = b(1)
      TimeStart = b(2)
      TimeEnd = b(3)
      yy1 = DatePart("yyyy", TimeEnd)
      mm1 = DatePart("m", TimeEnd)
      dd1 = DatePart("d", TimeEnd)
      hh1 = DatePart("h", TimeEnd)

      yy0 = DatePart("yyyy", TimeStart)
      mm0 = DatePart("m", TimeStart)
      dd0 = DatePart("d", TimeStart)
      hh0 = DatePart("h", TimeStart)
      
      Call yrsf(yy0, mm0, dd0, hh0, StartTime)
      Call yrsf(yy1, mm1, dd1, hh1, EndTime)
      LongTime = nd12(yy0, mm0, dd0, hh0, gltt, yy1, mm1, dd1, hh1) + 1
      glmm = nd12(yy0, mm0, dd0, hh0, gltm, yy1, mm1, dd1, hh1) + 1
glnn = LongTime
glnn1 = glnn
End If
b.MoveNext
Loop
b.Close
glmm = glnn
End Sub
Sub findtime()

Dim i As Integer, j As Integer

Dim it As Long, im As Integer, id As Integer
Dim yy0 As Integer, mm0 As Integer, dd0 As Integer, hh0 As Integer, nn0 As Integer
Dim yy1 As Integer, mm1 As Integer, dd1 As Integer, hh1 As Integer, nn1 As Integer

If b.state = 1 Then
  b.Close
Else
End If
bname = "dataflood" + dyly
b.CursorLocation = adUseClient
b.Open bname, cn

If Not b.BOF Then
b.MoveFirst
End If
Do While Not b.EOF
i = b(0)
If i = FloodNo(1, CountFlood) Then
      NoYear = b(1)
      TimeStart = b(2)
      TimeEnd = b(3)
      yy1 = DatePart("yyyy", TimeEnd)
      mm1 = DatePart("m", TimeEnd)
      dd1 = DatePart("d", TimeEnd)
      hh1 = DatePart("h", TimeEnd)
  
      yy0 = DatePart("yyyy", TimeStart)
      mm0 = DatePart("m", TimeStart)
      dd0 = DatePart("d", TimeStart)
      hh0 = DatePart("h", TimeStart)
     
      
      Call yrsf(yy0, mm0, dd0, hh0, StartTime)
      Call yrsf(yy1, mm1, dd1, hh1, EndTime)
      LongTime = nd12(yy0, mm0, dd0, hh0, gltt, yy1, mm1, dd1, hh1) + 1
      glmm = nd12(yy0, mm0, dd0, hh0, gltm, yy1, mm1, dd1, hh1) + 1
glnn = LongTime
glnn1 = glnn
End If
b.MoveNext
Loop
b.Close
glmm = glnn
End Sub
Sub findtimed(CountFlood As Integer)

Dim i As Integer, j As Integer
Dim b As New ADODB.Recordset

Dim it As Long, im As Integer, id As Integer
Dim yy0 As Integer, mm0 As Integer, dd0 As Integer, hh0 As Integer
Dim yy1 As Integer, mm1 As Integer, dd1 As Integer, hh1 As Integer

bname = "dayflood" + CStr(dyly)
b.CursorLocation = adUseClient
b.Open bname, cn

    
If Not b.BOF Then
  b.MoveFirst
End If
Do While Not b.EOF
i = b(0)
If i = FloodNo(1, CountFlood) Then
      
      NoYear = b(1)
      TimeStart = b(2)
      TimeEnd = b(3)
      yy1 = DatePart("yyyy", TimeEnd)
      mm1 = DatePart("m", TimeEnd)
      dd1 = DatePart("d", TimeEnd)
      hh1 = DatePart("h", TimeEnd)
    

      yy0 = DatePart("yyyy", TimeStart)
      mm0 = DatePart("m", TimeStart)
      dd0 = DatePart("d", TimeStart)
      hh0 = DatePart("h", TimeStart)
      
      
      Call yrsfd(yy0, mm0, dd0, StartTimeD)
      Call yrsfd(yy1, mm1, dd1, EndTimeD)
      LongTimeD = DateSerial(yy1, mm1, dd1) - DateSerial(yy0, mm0, dd0) + 1
      glnn = LongTimeD
      glnn1 = glnn
End If
b.MoveNext
Loop
b.Close

     NoYearr = yy1
End Sub
Sub findnumber()

Dim i As Integer, j As Integer, k As Integer
bname = "selecflood" + dyly
sql1 = "select * from " + bname + " Order by [No] "

rst.CursorLocation = adUseClient
rst.Open sql1, cn

NumberNo = rst.RecordCount
ReDim FloodNo(2, NumberNo)

If Not rst.BOF Then
rst.MoveFirst
For i = 1 To NumberNo
  FloodNo(1, i) = rst(0)
rst.MoveNext
Next i
End If
rst.Close

bname = "dataflood" + dyly
sql1 = "select * from " + bname + " Order by [No] "
rst.CursorLocation = adUseClient
rst.Open sql1, cn
k = rst.RecordCount

If Not rst.BOF Then
rst.MoveFirst
j = 1
For i = 1 To k
 If FloodNo(1, j) = rst(0) Then
    FloodNo(2, j) = rst(1)
    j = j + 1
    If j > NumberNo Then GoTo cc
  End If
rst.MoveNext
Next i
cc:
End If
rst.Close
End Sub
Sub findnumberd()

Dim i As Integer, j As Integer, k As Integer

bname = "selecyear" + Basin
sql1 = "select * from " + bname + " Order by [No] "
rst.CursorLocation = adUseClient
rst.Open sql1, cn

NumberNo = rst.RecordCount
ReDim FloodNo(2, NumberNo)

If Not rst.BOF Then
rst.MoveFirst
For i = 1 To NumberNo
  FloodNo(1, i) = rst(0)
rst.MoveNext
Next i
End If
rst.Close

bname = "dayflood" + CStr(dyly)
sql1 = "select * from " + bname + " Order by [No]  "
rst.CursorLocation = adUseClient
rst.Open sql1, cn
k = rst.RecordCount

If Not rst.BOF Then
rst.MoveFirst
j = 1
For i = 1 To k
 If FloodNo(1, j) = rst(0) Then
    FloodNo(2, j) = rst(1)
    j = j + 1
    If j > NumberNo Then GoTo cc
 End If
rst.MoveNext
Next i
cc:
End If
rst.Close
End Sub
Sub findshownumber()

Dim i As Integer, j As Integer, k As Integer
bname = "dataflood"
sql1 = "select * from " + bname + " Order by [No] "

rst.CursorLocation = adUseClient
rst.Open sql1, cn

NumberNo = rst.RecordCount
ReDim FloodNo(2, NumberNo)

If Not rst.BOF Then
rst.MoveFirst
For i = 1 To NumberNo
  FloodNo(1, i) = rst(0)
rst.MoveNext
Next i
End If
rst.Close

End Sub
Sub findbasind()

Dim i As Integer, j As Integer

bname = "selecbasin"

rst.CursorLocation = adUseClient
rst.Open bname, cn

If Not rst.BOF Then
rst.MoveFirst
j = rst(0)
BasinNa = j
End If
rst.Close

bname = "totalbasin"
b.CursorLocation = adUseClient
b.Open bname, cn
     
If Not b.BOF Then
b.MoveFirst
End If
Do While Not b.EOF
i = b(0)
If i = j Then
  dylyc = b(1)
  dyly = b(2)
End If
b.MoveNext
Loop
b.Close

End Sub
Sub findbasin()

Dim i As Integer, j As Integer

bname = "selecbasin"

rst.CursorLocation = adUseClient
rst.Open bname, cn

If Not rst.BOF Then
rst.MoveFirst
j = rst(0)
BasinNa = j
End If
rst.Close

bname = "totalbasin"
rst.CursorLocation = adUseClient
rst.Open bname, cn
    
If Not rst.BOF Then
  rst.MoveFirst
End If
Do While Not rst.EOF
i = rst(0)
If i = j Then
  dylyc = rst(1)
  dyly = rst(2)
End If
rst.MoveNext
Loop
rst.Close
Basin = dyly

End Sub

Sub showsnumber()

Dim i As Integer, j As Integer
bname = "showsflood"
rst.CursorLocation = adUseClient
rst.Open bname, cn

NumberNo = rst.RecordCount
ReDim FloodNo(2, NumberNo)
If Not rst.BOF Then
 rst.MoveFirst
   NoFlood = rst(0)
End If
rst.Close

End Sub
Sub showsnumberd()
Dim i As Integer, j As Integer
bname = "showsflood"
rst.CursorLocation = adUseClient
rst.Open bname, cn

NumberNo = rst.RecordCount
ReDim FloodNo(2, NumberNo)
If Not rst.BOF Then
 rst.MoveFirst
   NoFlood = rst(0)
End If
rst.Close

End Sub

Sub showstime(NoFlood)

Dim i As Integer, j As Integer
Dim yy0 As Integer, mm0 As Integer, dd0 As Integer, hh0 As Integer
Dim yy1 As Integer, mm1 As Integer, dd1 As Integer, hh1 As Integer

bname = "dataflood"
b.CursorLocation = adUseClient
b.Open bname, cn
    
If Not b.BOF Then
b.MoveFirst
End If
Do While Not b.EOF
i = b(0)
If i = NoFlood Then

      NoYear = b(1)
      TimeStart = b(2)
      TimeEnd = b(3)
      yy1 = DatePart("yyyy", TimeEnd)
      mm1 = DatePart("m", TimeEnd)
      dd1 = DatePart("d", TimeEnd)
      hh1 = DatePart("h", TimeEnd)

      yy0 = DatePart("yyyy", TimeStart)
      mm0 = DatePart("m", TimeStart)
      dd0 = DatePart("d", TimeStart)
      hh0 = DatePart("h", TimeStart)
      
      Call yrsf(yy0, mm0, dd0, hh0, StartTime)
      Call yrsf(yy1, mm1, dd1, hh1, EndTime)
      LongTime = nd12(yy0, mm0, dd0, hh0, gltt, yy1, mm1, dd1, hh1) + 1
      glnn = LongTime
      glnn1 = glnn
  End If
b.MoveNext
Loop
b.Close
End Sub
Sub showstimed(NoFlood)

Dim i As Integer, j As Integer

Dim yy0 As Integer, mm0 As Integer, dd0 As Integer, hh0 As Integer
Dim yy1 As Integer, mm1 As Integer, dd1 As Integer, hh1 As Integer


bname = "dayflood"
b.CursorLocation = adUseClient
b.Open bname, cn
    
If Not b.BOF Then
b.MoveFirst
End If
Do While Not b.EOF
i = b(0)
If i = NoFlood Then
       NoYear = b(1)
      TimeStart = b(2)
      TimeEnd = b(3)
      yy1 = DatePart("yyyy", TimeEnd)
      mm1 = DatePart("m", TimeEnd)
      dd1 = DatePart("d", TimeEnd)
      hh1 = DatePart("h", TimeEnd)

      yy0 = DatePart("yyyy", TimeStart)
      mm0 = DatePart("m", TimeStart)
      dd0 = DatePart("d", TimeStart)
      hh0 = DatePart("h", TimeStart)
      
      Call yrsfd(yy0, mm0, dd0, StartTimeD)
      Call yrsfd(yy1, mm1, dd1, EndTimeD)
      LongTimeD = DateSerial(yy1, mm1, dd1) - DateSerial(yy0, mm0, dd0) + 1
      glnn = LongTimeD
      glnn1 = glnn
End If
b.MoveNext
Loop
b.Close
glnn = LongTimeD
glnn1 = glnn
End Sub

Sub calcu_averagp_day(m As Integer)

Dim i As Integer, j As Integer, jj As Integer, it As Long, k As Integer, ii As Integer
Dim it1 As Long, it2 As Long
Dim R_name() As String, NoRS() As Integer, Na As Integer, MaxNoSt As Integer

MaxNoSt = 4

bname = "R_name_Averap"
sql1 = "select * from " + bname + " order by [±àºÅ]  "
rst.CursorLocation = adUseClient
rst.Open sql1, cn

Na = rst.RecordCount
ReDim R_name(rst.RecordCount, MaxNoSt), NoRS(Na)

For j = 1 To rst.RecordCount
     NoRS(j) = rst(1)
     For i = 1 To NoRS(j)
         R_name(j, i) = rst(i + 1)
     Next i
     rst.MoveNext
    Next j
rst.Close

Dim pp() As Single, p() As Single, Np() As Integer
ReDim pp(m, Na), p(2 * m, MaxNoSt), Np(MaxNoSt)
Dim pinjun As Single, Cname As String, NoR As Integer


For k = 1 To Na

 For i = 1 To m
   For j = 1 To MaxNoSt
      p(i, j) = 0
   Next j
 Next i

For jj = 1 To NoRS(k)
  it1 = gsdsj(1)
  it2 = gsdsj(m)
  bname = "rain_day"
  Cname = R_name(k, jj)
  sql1 = "select " + CStr(Cname) + " from " + bname + " where dt between  " + CStr(it1) + " and " + CStr(it2) + "  order by DT  "
  rst.CursorLocation = adUseClient
  rst.Open sql1, cn
  ii = rst.RecordCount
  For j = 1 To ii
           If rst(0) > 0.0001 Then
               p(j, jj) = rst(0)
           Else
               p(j, jj) = 0
           End If
           rst.MoveNext
  Next j
  rst.Close
Next jj

For jj = 1 To NoRS(k)
  Np(jj) = 0
  For j = 1 To m
     If p(j, jj) > 0 Then
        Np(jj) = 1
     Else
     End If
  Next j
Next jj

NoR = 0
For jj = 1 To NoRS(k)
  If Np(jj) = 1 Then
     NoR = NoR + 1
  Else
  End If
Next jj

For j = 1 To m
  pinjun = 0
  For i = 1 To NoRS(k)
      pinjun = pinjun + p(j, i) / (NoR + 0.0001)
  Next i
  pp(j, k) = pinjun
Next j
 
Next k
 
bname = "R_average_day"
rst.Open bname, cn, adOpenDynamic, adLockOptimistic
cn.Execute "Delete * from R_average_day"
j = 1
Do While j <= m
        rst.AddNew
        rst(0) = gsdsj(j)
        For i = 1 To Na
           rst(i) = Int(pp(j, i) * 100) / 100
        Next i
        rst.Update
       j = j + 1
Loop
rst.Close

End Sub
Sub calcu_pan_day()

Dim i As Integer, j As Integer, jj As Integer, it As Long, k As Integer
Dim it1 As Long, it2 As Long

Na = 15
Dim pp() As Single, p() As Single
ReDim pp(Na, glnn), p(Na, glday)

For i = 1 To Na
   For j = 1 To glday
      pp(i, j) = 0
   Next j
 Next i

  it1 = glchsdsj(1)
  it2 = glchsdsj(glnn)
  bname = "R_average_period"
  rst.CursorLocation = adUseClient
  rst.Open sql1, cn
 
  For j = 1 To glnn
    For i = 1 To Na
       If rst(i) > 0.0001 Then
          pp(i, j) = rst(i)
       Else
          pp(i, j) = 0
       End If
    Next i
      rst.MoveNext
  Next j
  rst.Close

Dim kk As Integer

For i = 1 To Na
  
  For j = 1 To glday
    
    p(i, j) = 0
    jj = (j - 1) * 24
    
    For k = 1 To 24
      kk = jj + k
      If kk <= glnn Then
        p(i, j) = p(i, j) + pp(i, kk)
      Else
        p(i, j) = p(i, j)
      End If
    Next k
  
  Next j

Next i
 
bname = "Rain_Day_pan_D"
rst.Open bname, cn, adOpenDynamic, adLockOptimistic
j = 1
Do While j <= glday
        rst.AddNew
        rst(0) = CountFlood
        rst(1) = sdsj(j)
        rst(2) = Ddata(j)
        For i = 1 To Na
           rst(i + 2) = Int(p(i, j) * 100) / 100
        Next i
        rst.Update
       j = j + 1
Loop
rst.Close

Dim PMax() As Single, imax() As Integer
ReDim PMax(Na), imax(Na)

For i = 1 To Na
  PMax(i) = 0
  For j = 1 To glday
    If p(i, j) > PMax(i) Then
       PMax(i) = p(i, j)
       k = j
    Else
    End If
  Next j
  imax(i) = k
Next i

bname = "Rain_Day_pan_Max"
rst.Open bname, cn, adOpenDynamic, adLockOptimistic
        rst.AddNew
        rst(0) = CountFlood
        rst(1) = sdsj(1)
        For i = 1 To Na
           rst(i + 1) = Int(PMax(i) * 100) / 100
        Next i
        rst.Update
rst.Close


bname = "Rain_Day_pan_25"
rst.Open bname, cn, adOpenDynamic, adLockOptimistic
j = 1
Do While j <= glday
        rst.AddNew
        rst(0) = CountFlood
        rst(1) = sdsj(j)
        rst(2) = Ddata(j)
        For i = 1 To Na
           If p(i, j) >= 25 Then
              rst(i + 2) = Int(p(i, j) * 100) / 100
           Else
              rst(i + 2) = Int(0 * 100) / 100
           End If
        Next i
        rst.Update
       j = j + 1
Loop
rst.Close

bname = "Rain_Day_pan_50"
rst.Open bname, cn, adOpenDynamic, adLockOptimistic
j = 1
Do While j <= glday
        rst.AddNew
        rst(0) = CountFlood
        rst(1) = sdsj(j)
        rst(2) = Ddata(j)
        For i = 1 To Na
           If p(i, j) >= 50 Then
              rst(i + 2) = Int(p(i, j) * 100) / 100
           Else
              rst(i + 2) = Int(0 * 100) / 100
           End If
        Next i
        rst.Update
       j = j + 1
Loop
rst.Close

bname = "Rain_Day_pan_100"
rst.Open bname, cn, adOpenDynamic, adLockOptimistic
j = 1
Do While j <= glday
        rst.AddNew
        rst(0) = CountFlood
        rst(1) = sdsj(j)
        rst(2) = Ddata(j)
        For i = 1 To Na
           If p(i, j) >= 100 Then
              rst(i + 2) = Int(p(i, j) * 100) / 100
           Else
              rst(i + 2) = Int(0 * 100) / 100
           End If
        Next i
        rst.Update
       j = j + 1
Loop
rst.Close

End Sub
Sub calcu_pan()

Dim i As Integer, j As Integer, jj As Integer, it As Long, k As Integer
Dim it1 As Long, it2 As Long, sql1 As String

Na = 15
Dim pp() As Single, p() As Single
ReDim pp(Na, glnn), p(Na, glday)

For i = 1 To Na
   For j = 1 To glday
      pp(i, j) = 0
   Next j
 Next i

  it1 = glchsdsj(1)
  it2 = glchsdsj(glnn)
  bname = "R_average_period"
  sql1 = "select * from " + bname + " where dt between  " + CStr(it1) + " and " + CStr(it2) + "  order by DT  "
  rst.CursorLocation = adUseClient
  rst.Open sql1, cn
 
  For j = 1 To glnn
    For i = 1 To Na
       If rst(i) > 0.0001 Then
          pp(i, j) = rst(i)
       Else
          pp(i, j) = 0
       End If
    Next i
      rst.MoveNext
  Next j
  rst.Close

Dim kk As Integer

For i = 1 To Na
  
  For j = 1 To glday
    
    p(i, j) = 0
    jj = (j - 1) * 6
    
    For k = 1 To 6
      kk = jj + k
      If kk <= glnn Then
        p(i, j) = p(i, j) + pp(i, kk)
      Else
        p(i, j) = p(i, j)
      End If
    Next k
  
  Next j

Next i
 
bname = "Rain_Day_pan_D"
rst.Open bname, cn, adOpenDynamic, adLockOptimistic
j = 1
Do While j <= glday
        rst.AddNew
        rst(0) = CountFlood
        rst(1) = Ddata(j)
        For i = 1 To Na
           rst(i + 1) = Int(p(i, j) * 100) / 100
        Next i
        rst.Update
       j = j + 1
Loop
rst.Close


End Sub
Sub findnumberd_day()

Dim i As Integer, j As Integer, k As Integer

bname = "Year_Data"
b.CursorLocation = adUseClient
b.Open bname, cn

NumberNo = b.RecordCount
b.Close
End Sub
Sub findtimed_day(CountFlood As Integer)

Dim i As Integer, j As Integer
Dim b As New ADODB.Recordset

Dim it As Long, im As Integer, id As Integer
Dim yy0 As Integer, mm0 As Integer, dd0 As Integer, hh0 As Integer
Dim yy1 As Integer, mm1 As Integer, dd1 As Integer, hh1 As Integer

bname = "Year_Data"
b.CursorLocation = adUseClient
b.Open bname, cn
    
If Not b.BOF Then
  b.MoveFirst
End If
Do While Not b.EOF
i = b(0)
If i = CountFlood Then
      
      NoYear = b(1)
      TimeStart = b(2)
      TimeEnd = b(3)
      yy1 = DatePart("yyyy", TimeEnd)
      mm1 = DatePart("m", TimeEnd)
      dd1 = DatePart("d", TimeEnd)
      hh1 = DatePart("h", TimeEnd)

      yy0 = DatePart("yyyy", TimeStart)
      mm0 = DatePart("m", TimeStart)
      dd0 = DatePart("d", TimeStart)
      hh0 = DatePart("h", TimeStart)
      
      Call yrsfd(yy0, mm0, dd0, StartTimeD)
      Call yrsfd(yy1, mm1, dd1, EndTimeD)
      LongTimeD = DateSerial(yy1, mm1, dd1) - DateSerial(yy0, mm0, dd0) + 1
      glnn = LongTimeD
      glnn1 = glnn
End If
b.MoveNext
Loop
b.Close

     NoYearr = yy1
End Sub
