Attribute VB_Name = "zmqbb"
Sub rishow()

   Call findnumberd
   

     CountFlood = 1
     Call findtimed(CountFlood)
     NoFlood = FloodNo(1, CountFlood)
     Call daytimed(LongTimeD)
      
End Sub
Sub cishow()

   Call findnumber
   

     CountFlood = 1
     Call findtime
     NoFlood = FloodNo(1, CountFlood)
     Call timeint
      
End Sub
Sub zmqbbb(CountFlood As Integer)

Dim itb As Long, itbb As Long, it1 As Long, it2 As Long, it3 As Long, it As Long
Dim iy As Integer, im   As Integer, id  As Integer, ih As Integer, iii As Integer
Dim yy1   As Integer, mm1  As Integer, dd1  As Integer, hh1 As Integer

Call cszcx(itb, CountFlood)

 it1 = StartTime
 Call ymdh(it1, iy, im, id, ih)
 Call tymd(iy, im, id, -1, yy1, mm1, dd1)
 Call yrsfd(yy1, mm1, dd1, it)
 EndTimeD = it
 Call yrsfd(yy1, 6, 1, it3)
 
 If glztz = 0 Or (glztz = 1 And itb < EndTimeD) Then
  it = EndTimeD
   Call ymd(it, iy, im, id)
   Call tymd(iy, im, id, -BeginDay, yy1, mm1, dd1)
  If yy1 < iy Then
     yy1 = iy
     mm1 = 1
     dd1 = 1
  End If
  Call yrsfd(yy1, mm1, dd1, it1)
  StartTimeD = it1
Else
  StartTimeD = itb
End If

If StartTimeD < EndTimeD Then

  Call ymd(StartTimeD, yy1, mm1, dd1)
  Call ymd(EndTimeD, iy, im, id)
  LongTimeD = DateSerial(iy, im, id) - DateSerial(yy1, mm1, dd1) + 1
  Call daytimed(LongTimeD)
  'Call calcu_averagp_day(glnn)
  Call chyubasW
 
 Else
 End If
End Sub
Sub cszcx(itb As Long, CountFlood As Integer)
Dim sql1 As String, yyear As Long
yyear = FloodNo(2, CountFlood)

bname = "dast" + dyly
sql1 = "select * from " + bname + " where year = " + CStr(yyear) + " ORDER BY dt desc"

rd1.CursorLocation = adUseClient
rd1.Open sql1, cn

If rd1.EOF And rd1.BOF Then
   If rd1.BOF Or rd1.EOF Then
     MsgBox "日模型状态计算结果不存在！"
   End If
     itb = 0
     glztz = 0
Else
     glztz = 1
     itb = rd1("dt")
End If
rd1.Close
End Sub

