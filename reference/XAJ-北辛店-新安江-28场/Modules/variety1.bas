Attribute VB_Name = "variety1"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub savebbo(wo() As Single, wj() As Single, pj() As Single, qin() As Single, qinh() As Single, qjy() As Single, qcal() As Single, qobs() As Single, _
            rr() As Single, rin As Single, robsy As Single, rcaly As Single, ce As Single, qom As Single, qcm As Single, _
            eqm As Single, iom As Integer, icm As Integer, iem As Integer, dc As Single, m As Integer, emy As Single, eky As Single, eey As Single, area As Single, rrcc As Single, ew() As Single, _
            dc2 As Single, dc3 As Single, dc4 As Single)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim ppz As Single, rrz As Single
Dim i As Integer, j As Integer, it1 As Long, sitd As String, it2 As Long, itd As Long, pp As Single
Dim iy1 As Integer, im1 As Integer, id1 As Integer, ih1 As Integer
Dim ryx As Single, rwc As Single, rhg As Integer
Dim rcali As Single

Dim yxqmax As Single, hgqmax As Integer, wcqmax As Single

bname = "ybresu" + dyly
sql1 = "delete  from " + bname + " where 洪水起始时间 = " + CStr(glchsdsj(1))
bb.CursorLocation = adUseClient
bb.Open bname, cn, adOpenDynamic, adLockOptimistic
cn.Execute (sql1)


j = 1
Do While j <= glnn
        bb.AddNew
        bb(0) = glchsdsj(1)
        bb(1) = glchsdsj(j)
        bb(2) = Hdate(j)
        bb(3) = Int(pj(j) * 100) / 100
        bb(4) = Int(wj(j) * 100) / 100
        bb(5) = Int(rr(j) * 100) / 100
        bb(6) = Int(qinh(j) * 100) / 100
        bb(7) = Int(qjy(j) * 100) / 100
        bb(8) = Int(qobs(j) * 100) / 100
        bb(9) = Int(qcal(j) * 100) / 100
        bb.Update
       j = j + 1
Loop
bb.Close


bname = "ybresu" + dyly + "1"
sql1 = "delete  from " + bname + " where 洪水起始时间 = " + CStr(glchsdsj(1))
bb.CursorLocation = adUseClient
bb.Open bname, cn, adOpenDynamic, adLockOptimistic
cn.Execute (sql1)

pp = 0
rcali = 0
For j = 1 To m
 pp = pp + pj(j)
 rcali = rcali + qin(j)
Next j
rcali = rcali * 3600 * gltt / 10 ^ 6
ryx = robsy * 0.2
If ryx >= 20 Then ryx = 20
If ryx <= 3 Then ryx = 3
rwc = robsy - rcaly
rhg = 0
If Abs(rwc) < ryx Then rhg = 1

yxqmax = qom * 0.2
hgqmax = 0
wcqmax = qcm - qom
If Abs(wcqmax) < yxqmax Then hgqmax = 1

bb.AddNew
bb(0) = glchsdsj(1)
bb(1) = Int(pp * 10) / 10
bb(2) = Int(eey * 10) / 10
bb(3) = Int(rrcc * 10) / 10
bb(4) = Int(robsy * 10000) / 10000
bb(5) = Int(rcaly * 10000) / 10000
bb(6) = Int((robsy - rcaly) * 100) / 100
bb(7) = Int(ryx * 100) / 100
bb(8) = rhg
bb(9) = Int(ce * 100) / 100
bb(10) = Int(qom * 10) / 10
bb(11) = Int(qcm * 10) / 10
bb(12) = hgqmax
bb(13) = Int(eqm * 10) / 10
bb(14) = Int(iem)
bb(15) = Round(dc, 2)
''''''''''''''''''''''''''''''''''''''
bb(16) = Round(dc2, 2)
bb(17) = Round(dc3, 2)
bb(18) = Round(dc4, 2)
'''''''''''''''''''''''''''''''''''''''

bb.Update
bb.Close

End Sub
Sub savebboD(wj() As Single, pj() As Single, qrc() As Single, qjy() As Single, qcal() As Single, qobs() As Single, _
            rr() As Single, rin As Single, robsy As Single, rcaly As Single, ce As Single, qom As Single, qcm As Single, _
            eqm As Single, iom As Integer, icm As Integer, iem As Integer, dc As Single, m As Integer, emy As Single, _
            eky As Single, eey As Single, rrcc As Single)

Dim ppz As Single, rrz As Single

Dim i As Integer, j As Integer, it1 As Long, sitd As String, it2 As Long, pp As Single
Dim iy1 As Integer, im1 As Integer, id1 As Integer, ih1 As Integer
Dim qobsum As Single, qcasum As Single, qse As Single
Dim rcali As Single, rqjy As Single



bname = "ybresuD" + dyly + "1"

sql1 = "delete  from " + bname + " where 年份 = " + CStr(NoYear)
rd8.CursorLocation = adUseClient
rd8.Open bname, cn, adOpenDynamic, adLockOptimistic
cn.Execute (sql1)

pp = o
rcali = 0
rqjy = 0
For j = 1 To m
  pp = pp + pj(j)
  rcali = rcali + qrc(j)
  rqjy = rqjy + qjy(j)
Next j

rcali = rcali * gltt * 3600 / 10 ^ 8
rqjy = rqjy * gltt * 3600 / 10 ^ 8
rd8.AddNew
rd8(0) = NoYear
rd8(1) = Int(pp * 10) / 10
rd8(2) = Int(eey * 10) / 10
rd8(3) = Int(robsy * 10000) / 10000
rd8(4) = Int(rcaly * 10000) / 10000
rd8(5) = Int(rqjy * 10000) / 10000
rd8(6) = Int(rcali * 10000) / 10000
rd8(7) = Int(ce * 100) / 100
rd8(8) = Int(qom * 10) / 10
rd8(9) = Int(qcm * 10) / 10
rd8(10) = Int(eqm * 10) / 10
rd8(11) = Int(iem)
rd8(12) = Int(dc * 100) / 100
rd8.Update
rd8.Close

bname = "ybresuD" + dyly + "0"

sql1 = "delete  from " + bname + " where 年份 = " + CStr(NoYear)
rd8.CursorLocation = adUseClient
rd8.Open bname, cn, adOpenDynamic, adLockOptimistic
cn.Execute (sql1)

qobsum = 0
qcasum = 0
For j = 90 To m - 92
  qobsum = qobsum + qobs(j) / (m - 90 - 92)
  qcasum = qcasum + qcal(j) / (m - 90 - 92)
Next j
qse = (qobsum - qcasum) / (qobsum + 0.001) * 100
rd8.AddNew
rd8(0) = NoYear
rd8(1) = Int(qobsum * 100) / 100
rd8(2) = Int(qcasum * 100) / 100
rd8(3) = Int(qse * 100) / 100
rd8.Update
rd8.Close

bname = "ybresuD" + dyly
it1 = sdsj(1)
it2 = sdsj(LongTime)
sql1 = "delete  from " + bname + " where 年份 = " + CStr(NoYear)

b.CursorLocation = adUseClient
b.Open bname, cn, adOpenDynamic, adLockOptimistic
cn.Execute (sql1)

j = 1
Do While j <= LongTimeD
        b.AddNew
        b(0) = NoYear
        b(1) = gsdsj(j)
        b(2) = Int(pj(j) * 100) / 100
        b(3) = Int(wj(j) * 100) / 100
        b(4) = Int(rr(j) * 100) / 100
        b(5) = Int(qrc(j) * 100) / 100
        b(6) = Int(qjy(j) * 100) / 100
        b(7) = Int(qobs(j) * 100) / 100
        b(8) = Int(qcal(j) * 100) / 100
        b.Update
       j = j + 1
Loop
b.Close



End Sub
