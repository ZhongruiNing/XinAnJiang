Attribute VB_Name = "jyanybfa"
Sub jyanyb()
Dim pa As Single, pj() As Single, r1() As Single, r2() As Single, qcal() As Single, qobs() As Single
Dim i As Integer, j As Integer, u As Single, k As Integer
Dim it As Long, ite As Long, area As Single
Dim iy As Integer, im As Integer, id As Integer, ih As Integer
Dim yy1 As Integer, mm1 As Integer, dd1 As Integer, hh1 As Integer
Dim fa() As Single, nn() As Integer, xx() As Single, qout() As Single

Dim qc() As Single, p() As Single, pf() As Single, rf() As Single, qads() As Single, pp() As Single, qcc() As Single

Call inputparapa(fa, nn, xx)

ReDim p(glnn, Na), pjj(Na, glnn), pf(Na, glnn), pj(glnn), rf(Na, glnn), qc(glnn), qads(glnn), pp(glnn), qcal(glnn), qcc(Na, glnn), qout(glnn)

Call inputrainJY(p)
  
For k = 1 To Na
  For j = 1 To glnn
     pf(k, j) = p(j, k)
  Next j
Next k
 
For k = 1 To Na
  pjj(k, 0) = 0
  For j = 1 To glnn
      pjj(k, j) = pjj(k, j - 1) + pf(k, j)
  Next j
Next k

it = glchsdsj(1)
Call ymdh(it, iy, im, id, ih)
Call tymd(iy, im, id, -1, yy1, mm1, dd1)
Call yrsfd(yy1, mm1, dd1, ite)

bname = "dastlspa"
sql1 = "select [Pa] from " + bname + "  where dt =" + CStr(ite)
b.CursorLocation = adUseClient
b.Open sql1, cn, adOpenStatic

If Not b.BOF Then
 Else
   MsgBox "没有当天的Pa，没有计算，按键退出！"
   Exit Sub
End If
  pa = b(0)
b.Close

For k = 1 To Na

  For i = 1 To glnn
    pj(i) = pjj(k, i)
  Next i

  Call chanliu(pa, pj, r1, r2)
   
   For i = 1 To glnn
    rf(k, i) = r2(i)
  Next i

Next k

For i = 1 To glnn
    qcal(i) = 0
Next i

area = 0
For k = 1 To Na
  
  area = area + fa(k)
  For i = 1 To glnn
    r2(i) = rf(k, i)
  Next i
  
  If dyly = "gx" Then
      u = fa(k) / gltt / 3.6
      Call uhqc(glnn, r2, u, Unit(k + 6), glnn, qc)
  Else
      
      u = fa(k) / gltt / 3.6
    If jy = 3 Then
      Call cslinear(r2, u, qc)
    Else
      Call uhqc(glnn, r2, u, Unit(k), glnn, qc)
    End If
    
  End If
  
  If jy = 3 Or jy = 2 Then
     For i = 1 To glnn
       qcc(k, i) = qc(i)
     Next i
  Else
  End If
     
  If jy = 1 Then
    For i = 1 To glnn
      qcal(i) = qcal(i) + qc(i)
    Next i
  Else
  End If
Next k

If jy = 2 Or jy = 3 Then
  Call JYqinflow(Na, qcc, qout, nn, xx)
Else
End If


Dim qrc() As Single, ia As Integer
ia = 1
Call qinflowD(ia, qc, qrc)

If jy = 2 Or jy = 3 Then
  For i = 1 To glnn
     qcal(i) = qout(i) + qrc(i) + Qjiliu
  Next i
Else
End If


If jy = 1 Then
   For i = 1 To glnn
     qcal(i) = qcal(i) + qrc(i) + Qjiliu
   Next i
Else
End If

ReDim qobs(glnn)
Call inputq(qobs)

For i = 1 To glnn
  pj(i) = 0
  pp(i) = 0
  r2(i) = 0
  For k = 1 To Na
    pp(i) = pp(i) + pf(k, i) / Na
    pj(i) = pj(i) + pjj(k, i) / Na
    r2(i) = r2(i) + rf(k, i) / Na
  Next k
Next i

r1(0) = 0
For i = 1 To glnn
   r1(i) = r1(i - 1) + r2(i)
Next i
Dim ce As Single, qom As Single, qcm As Single, eqm As Single
Dim iom As Integer, icm As Integer, iem As Integer, dc As Single
Dim robsy As Single, rcaly As Single, rin As Single
Dim wc() As Single, wo() As Single, ew() As Single
ReDim wc(5), wo(5), ew(5)

Call Charact_Watershed_jy_Q(qcal, qobs, robsy, rcaly, ce, qom, qcm, eqm, iom, icm, iem, dc, area)

Call Charact_Watershed_W(qcal, icm, wc)

Call Charact_Watershed_W(qobs, iom, wo)

For i = 1 To 5
  ew(i) = (wo(i) - wc(i)) / (wo(i) + 0.0001)
Next i

Call savejy(pj, r1, r2, pp, qcal, qobs, qads, qrc, robsy, rcaly, ce, qom, qcm, eqm, iom, icm, iem, dc, area, ew)
End Sub
Sub uhqc(mi As Integer, rd() As Single, u As Single, bb As String, mo As Integer, qc() As Single)
Dim i As Integer
Dim j As Integer
Dim ij As Integer
Dim mu As Integer, uh() As Single
Call readUnit(mu, uh, bb)
ReDim qc(glnn + mu + 1)
For i = 1 To mo
qc(i) = 0
Next i
For i = 1 To mi
  For j = 1 To mu
      ij = i + j - 1
      qc(ij) = qc(ij) + rd(i) * uh(j) * u
  Next j
Next i

End Sub
Sub cslinear(rd() As Single, u As Single, qc() As Single)
Dim i As Integer
Dim j As Integer
Dim ij As Integer
ReDim qc(glnn + 1)
cs = 0.01 ^ (gltt / 24#)
For i = 1 To glnn
qc(i) = 0
Next i

For i = 1 To glnn
   qc(i) = cs * qc(i - 1) + rd(i) * (1 - cs) * u
Next i

End Sub

Sub chanliu(pa As Single, pj() As Single, r1() As Single, r2() As Single)
Dim i As Integer, j As Integer, p As Single, nn As Integer

Dim pr() As Single, rr() As Single, pr1() As Single, pr2() As Single, rr1() As Single, rr2() As Single
Dim rp1 As Single, rp2 As Single, pa1 As Single, pa2 As Single

Call readPRcur(pa, pr, rr, nn, pa1, pa2)
ReDim r1(glnn), pp(glnn), r2(glnn), pr1(nn), pr2(nn), rr1(nn), rr2(nn)
 
 For i = 1 To nn
    pr1(i) = pr(1, i)
    pr2(i) = pr(2, i)
    rr1(i) = rr(1, i)
    rr2(i) = rr(2, i)
 Next i
 
    pj(0) = 0
    For i = 0 To glnn
       p = pj(i)
       Call pwxcz(nn, pr1, rr1, p, rp1)
       Call pwxcz(nn, pr2, rr2, p, rp2)
       r1(i) = rp2 + (rp1 - rp2) / (pa1 - pa2) * (pa - pa2)
    Next i
    For i = 1 To glnn
        r2(i) = r1(i) - r1(i - 1)
     If r2(i) <= 0 Then
        r2(i) = 0
     Else
     End If
    Next i

End Sub
Sub readUnit(mu As Integer, uh() As Single, bb As String)
Dim i As Integer

bname = "Unit_Gx"
sql1 = "select " + CStr(bb) + "  from " + bname + " order by [序号] "
b.CursorLocation = adUseClient
b.Open sql1, cn
b.MoveLast
mu = b.RecordCount
b.MoveFirst
ReDim uh(mu)
For j = 1 To mu
     uh(j) = b(0)
     b.MoveNext
Next j
b.Close
   
End Sub
Sub readPRcur(pa As Single, pr() As Single, rr() As Single, nn As Integer, pa1 As Single, pa2 As Single)
Dim i As Integer

bname = "PPa_R"
sql1 = "select *  from " + bname + " order by [编号] "
b.CursorLocation = adUseClient
b.Open sql1, cn
nn = b.RecordCount
ReDim pr(2, nn), rr(2, nn)

If pa <= 10 Then
  pa1 = 0
  pa2 = 10
  For j = 1 To nn
      pr(1, j) = b(1)
      rr(1, j) = b(2)
      pr(2, j) = b(3)
      rr(2, j) = b(4)
      b.MoveNext
  Next j
Else
End If

If pa <= 20 And pa > 10 Then
    pa1 = 10
    pa2 = 20

   For j = 1 To nn
      pr(1, j) = b(3)
      rr(1, j) = b(4)
      pr(2, j) = b(5)
      rr(2, j) = b(6)
      b.MoveNext
   Next j
Else
End If

If pa <= 30 And pa > 20 Then
  pa1 = 20
  pa2 = 30

  For j = 1 To nn
      pr(1, j) = b(5)
      rr(1, j) = b(6)
      pr(2, j) = b(7)
      rr(2, j) = b(8)
      b.MoveNext
  Next j
Else
End If

If pa <= 40 And pa > 30 Then
  pa1 = 30
  pa2 = 40

  For j = 1 To nn
      pr(1, j) = b(7)
      rr(1, j) = b(8)
      pr(2, j) = b(9)
      rr(2, j) = b(10)
      b.MoveNext
  Next j

Else
End If

If pa <= 50 And pa > 40 Then
  pa1 = 40
  pa2 = 50

  For j = 1 To nn
      pr(1, j) = b(9)
      rr(1, j) = b(10)
      pr(2, j) = b(11)
      rr(2, j) = b(12)
      b.MoveNext
  Next j
Else
End If

If pa <= 60 And pa > 50 Then
   pa1 = 50
   pa2 = 60

  For j = 1 To nn
      pr(1, j) = b(11)
      rr(1, j) = b(12)
      pr(2, j) = b(13)
      rr(2, j) = b(14)
      b.MoveNext
  Next j
Else
End If

If pa <= 70 And pa > 60 Then
    pa1 = 60
    pa2 = 70

  For j = 1 To nn
      pr(1, j) = b(13)
      rr(1, j) = b(14)
      pr(2, j) = b(15)
      rr(2, j) = b(16)
      b.MoveNext
  Next j
Else
End If

b.Close

End Sub
Sub pwxcz(N As Integer, x() As Single, Y() As Single, u As Single, f As Single)
Dim i As Single
Dim l As Single
Dim x1 As Single, x2 As Single, x3 As Single
Dim a1 As Single, a2 As Single, a3 As Single
l = N - 1
If u < x(1) Then
   f = 0
Else
  For i = 2 To l
    If u <= x(i) Then GoTo 20
  Next i
    i = l
    GoTo 30
20:
   If i = 2 Then GoTo 30
   If u - x(i - 1) <= x(i) - u Then i = i - 1
30:
   x1 = x(i - 1)
   x2 = x(i)
   x3 = x(i + 1)
   a1 = (u - x2) * (u - x3) / ((x1 - x2) * (x1 - x3))
   a2 = (u - x3) * (u - x1) / ((x2 - x3) * (x2 - x1))
   a3 = (u - x1) * (u - x2) / ((x3 - x1) * (x3 - x2))
   f = a1 * Y(i - 1) + a2 * Y(i) + a3 * Y(i + 1)
End If
End Sub
Sub savejy(pj() As Single, r1() As Single, r2() As Single, pp() As Single, qcal() As Single, qobs() As Single, qads() As Single, _
           qrc() As Single, robsy As Single, rcaly As Single, ce As Single, qom As Single, qcm As Single, _
            eqm As Single, iom As Integer, icm As Integer, iem As Integer, dc As Single, area As Single, ew() As Single)
            
Dim i As Integer, j As Integer, it1 As Long, sitd As String, it2 As Long, itd As Long, ppp As Single
Dim iy1 As Integer, im1 As Integer, id1 As Integer, ih1 As Integer

Dim rcali As Single

bname = "jyresu" + dyly + "1"
sql1 = "delete  from " + bname + " where 洪号 = " + CStr(NoFlood)
rd8.CursorLocation = adUseClient
rd8.Open bname, cn, adOpenDynamic, adLockOptimistic
cn.Execute (sql1)

ppp = pj(glnn1)
rcali = 0

For j = 1 To glnn
 rcali = rcali + qrc(j)
Next j
rcali = rcali * 3600 * gltt / 10 ^ 6

itd = glchsdsj(iom)

If dc <= 0 Then
  dc = 0
Else
End If

If Abs(ce) >= 200# Then
  ce = 0
Else
End If


rd8.AddNew
rd8(0) = NoFlood
rd8(1) = StartTime
rd8(2) = Int(ppp * 10) / 10
rd8(3) = Int((r1(glnn)) * 100) / 100
rd8(4) = Int((r1(glnn) * area / 1000) * 100) / 100
rd8(5) = Int(rcaly * 10000) / 10000
rd8(6) = Int(robsy * 10000) / 10000
rd8(7) = Int(rcali * 10000) / 10000
rd8(8) = Int(ce * 100) / 100
rd8(9) = Int(qom * 10) / 10
rd8(10) = glchsdsj(iom)
rd8(11) = Int(qcm * 10) / 10
rd8(12) = glchsdsj(icm)
rd8(13) = Int((qom - qcm) / qom * 100)
rd8(14) = iom - icm
rd8(15) = Int(dc * 100) / 100
rd8(16) = Int(iem)
rd8(17) = Int(ce * 100) / 100
rd8(18) = Int(ew(1) * 100)
rd8(19) = Int(ew(2) * 100)
rd8(20) = Int(ew(3) * 100)
rd8(21) = Int(ew(4) * 100)
rd8(22) = Int(ew(5) * 100)
rd8.Update
rd8.Close

bname = "jyresu" + dyly
sql1 = "delete  from " + bname + " where 洪号 = " + CStr(NoFlood)
b.CursorLocation = adUseClient
b.Open bname, cn, adOpenDynamic, adLockOptimistic
cn.Execute (sql1)


j = 1
Do While j <= glnn
        b.AddNew
        b(0) = NoFlood
        b(1) = glchsdsj(j)
        b(2) = Int(pj(j) * 100) / 100
        b(3) = Int(r1(j) * 100) / 100
        b(4) = Int(pp(j) * 100) / 100
        b(5) = Int(r2(j) * 100) / 100
        b(6) = Int(qobs(j) * 100) / 100
        b(7) = Int(qcal(j) * 100) / 100
        b(8) = Int(qrc(j) * 100) / 100
        b.Update
       j = j + 1
Loop
b.Close


End Sub
Sub UnitTransfer()
Dim i As Integer, j As Integer, k As Integer, mu As Integer, mm As Integer, j1 As Integer, j2 As Integer

Dim u1() As Single, u2() As Single, s() As Single, s1() As Single
Dim x1 As Single, x2 As Single, x3 As Single, y1 As Single, y2 As Single, y3 As Single
Dim x As Single, Y As Single

ReDim u1(18), u2(6, 100), s(100), s1(100)

mu = 17
mm = 34

For k = 1 To 6
  bname = "Unit"
  sql1 = "select  * from " + bname + " order by [序号] "
  b.CursorLocation = adUseClient
  b.Open sql1, cn
  mu = b.RecordCount

  For j = 0 To mu - 1
       u1(j) = b(k)
     b.MoveNext
  Next j
  b.Close
  
  For i = 0 To mu
    s(i) = 0
    For j = 0 To i
      s(i) = s(i) + u1(j)
    Next j
  Next i
  
  For i = mu + 1 To mm
    s(i) = s(mu)
  Next i
  
  For i = 0 To mm
    s1(i * 2) = s(i)
  Next i
   
  For i = 0 To mu
    x1 = i * 2
    x2 = (i + 1) * 2
    x3 = (i + 2) * 2
    y1 = s1(x1)
    y2 = s1(x2)
    y3 = s1(x3)
    x = x1 + 1
    Call No_linear_neicha(x1, y1, x2, y2, x3, y3, x, Y)
    s1(x) = Y
    x = x2 + 1
    Call No_linear_neicha(x1, y1, x2, y2, x3, y3, x, Y)
    s1(x) = Y
  Next i

  For i = 1 To mm
    u2(k, i) = (s1(i) - s1(i - 1))
  Next i
  
Next k

Open pathc + "\out.dat" For Output As #1
For j = 0 To mm
  Print #1, j, Int(u2(1, j) * 1000) / 1000, Int(u2(2, j) * 1000) / 1000, Int(u2(3, j) * 1000) / 1000, Int(u2(4, j) * 1000) / 1000, Int(u2(5, j) * 1000) / 1000, Int(u2(6, j) * 1000) / 1000
Next j
Close #1
End Sub
Sub No_linear_neicha(x1 As Single, y1 As Single, x2 As Single, y2 As Single, x3 As Single, y3 As Single, x As Single, Y As Single)
Dim i As Single
Dim l As Single, u As Single, f As Single
Dim a1 As Single, a2 As Single, a3 As Single
u = x
a1 = (u - x2) * (u - x3) / ((x1 - x2) * (x1 - x3))
a2 = (u - x3) * (u - x1) / ((x2 - x3) * (x2 - x1))
a3 = (u - x1) * (u - x2) / ((x3 - x1) * (x3 - x2))
f = a1 * y1 + a2 * y2 + a3 * y3
Y = f
End Sub
 Sub Charact_Watershed_jy_Q(qcal() As Single, qobs() As Single, robsy As Single, rcaly As Single, _
      ce As Single, qom As Single, qcm As Single, eqm As Single, iom As Integer, icm As Integer, _
            iem As Integer, dc As Single, area As Single)

 'input variables are qcal(j), qobs(j),j=1 to m
 'output variables are 实测洪量 robsy,计算洪量 rcaly(万方),相对误差 ce _
                       实测峰值 qom, 计算峰值 qcm, 相对误差 eqm _
                       实测峰时 iom, 计算峰时icm,峰现时间预报误差 iem(小时) _
                       确定性系数 dc
 'input variables are tt,min1
 Dim rqo As Single, rqc As Single, eqobs As Single, f0 As Single, fn As Single
 Dim j As Integer, min1 As Single, tt As Single, N As Integer
     N = glnn1
     tt = gltt
     m = glnn
     min1 = 0.000001
      rqo = 0#
      rqc = 0#
      For j = 1 To m
        rqo = rqo + qobs(j)
        rqc = rqc + qcal(j)
      Next j
      robsy = rqo * 3600 * tt / 10 ^ 6
      rcaly = rqc * 3600 * tt / 10 ^ 6
      ce = (robsy - rcaly) / (robsy + min1) * 100#
      eqobs = rqo / (m + min1)
      f0 = 0#
      fn = 0#
      For j = 1 To N
         f0 = f0 + (qobs(j) - eqobs) * (qobs(j) - eqobs)
         fn = fn + (qcal(j) - qobs(j)) * (qcal(j) - qobs(j))
      Next j
      f0 = f0 / (N + min1)
      fn = fn / (N + min1)
      dc = 1# - fn / (f0 + min1)
      qom = qobs(1)
      For j = 1 To m
         If qobs(j) > qom Then
           qom = qobs(j)
           iom = j
         End If
      Next j
      qcm = qcal(1)
      For j = 1 To m
          If qcal(j) > qcm Then
          qcm = qcal(j)
          icm = j
          End If
      Next j
      eqm = (qom - qcm) / (qom + min1) * 100#
      iem = (iom - icm)
  End Sub

