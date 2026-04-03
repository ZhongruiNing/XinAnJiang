Attribute VB_Name = "xajmd"
Sub rmxjss()

   Call findnumberd
   

  For CountFlood = 1 To NumberNo
     Call findtimed(CountFlood)
     NoFlood = FloodNo(1, CountFlood)
     Call daytimed(LongTimeD)
     Call chyubasD
    MDImain.StatusBar1.Panels(1) = "共有" & CStr(NumberNo) & "场洪水,现在计算到第" & CStr(CountFlood) & "场"
    DoEvents
  Next CountFlood
  'MsgBox "日模洪水计算完毕！", vbInformation, "提示信息"
  Call rishow
  yb12.Show
  Beep
End Sub
Sub chyubasW()

Dim ia As Integer, j As Integer, i As Integer
Dim qobs() As Single, qc() As Single, qrc() As Single, state() As Single, zdylp() As Single, qjy() As Single
Dim evap() As Single, InflwName() As String, m As Integer, rr() As Single, wj() As Single
Dim emy As Single, eky As Single, eey As Single, rrcc As Single, area As Single

Dim ce As Single, qom As Single, qcm As Single, eqm As Single
Dim iom As Integer, icm As Integer, iem As Integer, dc As Single
Dim robsy As Single, rcaly As Single, rin As Single
Dim pj() As Single, qcal() As Single

m = LongTimeD
Call inputparad(para1, para2, para3, Na, ia)

ReDim evap(m + Int(24 / gltt))
Call eevapD(evap)

ReDim zdylp(glnn, Na), state(Na, 7)

Call inputrainD(Na, zdylp, para3)
Call statehD(Na, state)
ReDim qobs(glnn + 10), qjy(glnn)
Call chmybD(para1, para2, evap, pj, qcal, qobs, zdylp, state, rr, Na, ia, wj, emy, eky, eey, rrcc, area)

End Sub
Sub chyubasD()

Dim ia As Integer, j As Integer, i As Integer
Dim qobs() As Single, qc() As Single, qrc() As Single, state() As Single, zdylp() As Single, qjy() As Single
Dim evap() As Single, InflwName() As String, m As Integer, rr() As Single, wj() As Single
Dim emy As Single, eky As Single, eey As Single, rrcc As Single, area As Single

Dim ce As Single, qom As Single, qcm As Single, eqm As Single
Dim iom As Integer, icm As Integer, iem As Integer, dc As Single
Dim robsy As Single, rcaly As Single, rin As Single
Dim pj() As Single, qcal() As Single

m = LongTimeD
Call inputparad(para1, para2, para3, Na, ia)

ReDim evap(m)
Call eevapD(evap)

ReDim zdylp(glnn, Na), state(Na, 7)

Call inputrainD(Na, zdylp, para3)
Call statehD(Na, state)
ReDim qobs(glnn + 10), qjy(glnn)
Call inputqD(qobs)
Call chmybD(para1, para2, evap, pj, qcal, qobs, zdylp, state, rr, Na, ia, wj, emy, eky, eey, rrcc, area)
Call qinflowD(ia, qc, qrc)
For i = 1 To m
  qjy(i) = qcal(i)
  qcal(i) = qcal(i) + qrc(i)
Next i
   Call Charact_Watershed_QD(qcal, qobs, qrc, m, rin, robsy, rcaly, ce, qom, qcm, eqm, iom, icm, iem, dc, area)

Call savebboD(wj, pj, qrc, qjy, qcal, qobs, rr, rin, robsy, rcaly, ce, qom, qcm, eqm, iom, icm, iem, dc, m, emy, eky, eey, rrcc)

End Sub
Sub chmybD(para1() As Single, para2() As Single, evap() As Single, _
       pj() As Single, qcal() As Single, qobs() As Single, zdylp() As Single, _
       state() As Single, rr() As Single, Na As Integer, ia As Integer, wj() As Single, _
       emy As Single, eky As Single, eey As Single, rrcc As Single, area As Single)
    
Dim k As Single, im As Single, b As Single, wum As Single, wlm As Single, _
    wm As Single, c As Single, sm As Single, ex As Single, kg As Single, _
    ki As Single, cg As Single, ci As Single, x As Single, kk As Single, _
    cs As Single, f As Single, tt As Single, k1 As Single, k2 As Single
    
Dim ttt As String, itt As Integer, cs1 As Single, lhh As Long, lag As Integer
Dim wdm As Single, wmm As Single, smm As Single, c0 As Single, c1 As Single, _
    c2 As Single, div As Single, min1 As Single, minn As Single, ly As Integer

Dim fp() As Single, mp() As Integer, cp() As Single
Dim wp() As Single, wup() As Single, wlp() As Single, sp() As Single, frp() As Single, _
    qsp() As Single, qip() As Single, qgp() As Single, lp As Integer
Dim numsq As Integer, qq() As Single, zz() As Single, numq As Integer

Dim m As Integer, iy2 As Integer, im2 As Integer, _
    id2 As Integer, ih2 As Integer, iy1 As Integer, im1 As Integer, id1 As Integer, _
    ih1 As Integer, itd As Long, itd1 As Long, itd2 As Long, sitd As String

Dim wdp() As Single, _
    ep() As Single, eup() As Single, elp() As Single, edp() As Single, _
     rp() As Single, rsp() As Single, _
     rip() As Single, rgp() As Single, _
    qxs() As Single, qxi() As Single, qxg() As Single
    
Dim w As Single, wu As Single, wl As Single, wd As Single, e As Single, _
    eu As Single, el As Single, ed As Single, fr As Single, s As Single, _
    r As Single, rs As Single, ri As Single, rg As Single, qqs As Single, _
    qqi As Single, qqg As Single
Dim pe As Single, nd As Integer, rd() As Single, ped() As Single, pedf() As Single, _
    qss As Single, qii As Single, qgg As Single, rq As Single, qx() As Single

Dim ww As Single, wwu As Single, wwl As Single, wwd As Single, em As Single, ek As Single, _
    ee As Single, rrs As Single, rri As Single, rrg As Single, p As Single, _
    ss As Single, ffr As Single, pp() As Single, qin() As Single, qres() As Single
Dim qs() As Single, qi() As Single, qg() As Single

Dim qsz As Single, qiz As Single, qgz As Single

Dim py As Single, rsy As Single, riy As Single
Dim rgy As Single, riny As Single, rresy As Single
Dim irs As Single, iia As Single, inn As Single, ifc As Single, imf As Single, ef As Single, ib As Single

Dim mp1 As Integer, mpmax As Integer, bb1 As Single, bb2 As Single, yy4 As Integer, _
    l As Integer, i As Integer, j As Integer, ct As Single, bb3(10) As Single, _
    bb4 As Variant, jj As Integer, i1 As Integer, i2 As Integer
Dim ke As Single

k = para1(1): b = para1(2): c = para1(3): wm = para1(4)
wum = para1(5): wlm = para1(6): im = para1(7): sm = para1(8)
ex = para1(9): kg = para1(10): ki = para1(11): cg = para1(12)
ci = para1(13): x = para1(14): tt = para1(15): cs = para1(16)
f = para1(17): lp = para1(18): ia = para1(19): lag = para1(20)
kp = para1(23)
area = f

ReDim mp(lp)
ReDim fp(lp), mp(lp), cp(lp), pp(lp), wp(lp), wlp(lp), wup(lp), qsp(lp), qip(lp), qgp(lp)
ReDim ep(lp), eup(lp), elp(lp), edp(lp), rp(lp), rsp(lp), rip(lp), rgp(lp), wdp(lp)
ReDim frp(lp), sp(lp)

For i = 1 To lp
   fp(i) = para2(1, i)
   mp(i) = para2(2, i)
Next i
mpmax = 1
For i = 1 To lp
jj = mp(i)
If jj > mpmax Then
mpmax = mp(i)
End If
Next i
mp1 = mpmax + 1
ReDim qxs(lp, mp1)
ReDim qxi(lp, mp1)
ReDim qxg(lp, mp1)
ReDim qx(mp1)
ttt = CStr(Int(tt))
ttt = (ttt)
itt = Int(tt)

wdm = wm - wum - wlm
wmm = (1# + b) * wm / (1# - im)
smm = (1# + ex) * sm
bb1 = kg + ki
bb2 = kg / ki
ki = (1# - (1# - bb1) ^ (tt / 24#)) / (1 + bb2)
kg = ki * bb2
cg = cg ^ (tt / 24#)
ci = ci ^ (tt / 24#)
cs = cs ^ (tt / 24#)
kk = tt
c0 = (0.5 * tt - kk * x) / (kk - kk * x + 0.5 * tt)
c1 = (kk * x + 0.5 * tt) / (kk - kk * x + 0.5 * tt)
c2 = (kk - kk * x - 0.5 * tt) / (kk - kk * x + 0.5 * tt)
For i = 1 To lp
  cp(i) = fp(i) * f / tt / 3.6
Next i

div = 8#
minn = 0.0001

m = glnn

  iia = 0.4
  inn = 0.5
  ifc = 10#
  imf = 0#
  ef = 0.05
  ib = 1#
  imf = 0
  ib = 0

'Open "a.txt" For Append As #2
Dim it1 As Long, it2 As Long

bname = "dast" + dyly
it1 = gsdsj(1)
it2 = gsdsj(LongTimeD)
sql1 = "delete * from " + bname + " where dt between  " + CStr(it1) + " and " + CStr(it2)
rdd.CursorLocation = adUseClient
rdd.Open bname, cn, adOpenDynamic, adLockOptimistic
cn.Execute (sql1)

ReDim qs(glnn), qi(glnn), qg(glnn), qsig(glnn + lag), qcal(glnn + lag)
ReDim pj(glnn), rr(glnn), wj(glnn)
 
 
 
      qsig(0) = qobs(1)
      qcal(0) = qobs(1)
   For i = 0 To lag
      qsig(i) = qobs(1)
      qcal(i) = qobs(1)
   Next i
 
 
 
 
 
 For i = 0 To lag
   qsig(i) = qobs(1)
 Next i
       
 If glztz = 0 Then

      qcal(0) = qobs(1)
      qs(0) = qcal(0) / 3#
      qi(0) = qcal(0) / 3#
      qg(0) = qcal(0) / 3#
 
      For i = 1 To lp
        wp(i) = state(0, 1)
        wup(i) = state(0, 2)
        wlp(i) = state(0, 3)
        wdp(i) = wp(i) - wup(i) - wlp(i)
        sp(i) = state(0, 5)
        frp(i) = state(0, 6)
        qsp(i) = qcal(0) / lp / 3#
        qip(i) = qcal(0) / lp / 3#
        qgp(i) = qcal(0) / lp / 3#
       Next i
       For i = 1 To lp
       For j = 1 To mp(i) + 1
            qxs(i, j) = qsp(i)
            qxi(i, j) = qip(i)
            qxg(i, j) = qgp(i)
          Next j
       Next i
   Else
   End If

  If glztz = 1 Then

      qcal(0) = qobs(1)
      qs(0) = qcal(0) / 3#
      qi(0) = qcal(0) / 3#
      qg(0) = qcal(0) / 3#
 
      For i = 1 To lp
        wp(i) = state(i, 1)
        wup(i) = state(i, 2)
        wlp(i) = state(i, 3)
        wdp(i) = wp(i) - wup(i) - wlp(i)
        sp(i) = state(i, 6)
        frp(i) = state(i, 7)
        qsp(i) = qcal(0) / lp / 3#
        qip(i) = qcal(0) / lp / 3#
        qgp(i) = qcal(0) / lp / 3#
    
      Next i

       For i = 1 To lp
       For j = 1 To mp(i) + 1
            qxs(i, j) = qsp(i)
            qxi(i, j) = qip(i)
            qxg(i, j) = qgp(i)
          Next j
       Next i
   Else
   End If
  
      py = 0
      emy = 0
      eky = 0
      eey = 0
      rsy = 0
      riy = 0
      rgy = 0
      rrcc = 0
      ke = 1
      
 For j = 1 To m
 
    If Ma <= 1 Then
    Else
      For ii = 1 To Ma
        If j >= Int(para4(ii, 1)) And j <= Int(para4(ii, 2)) Then
           k = para4(ii, 3)
           Exit For
        Else
        End If
      Next ii
    
    End If

    
    For i = 1 To lp
          pp(i) = zdylp(j, i) * ke
      Next i
      em = evap(j)
      ek = k * em
      p = 0
      ww = 0
      wwu = 0
      wwl = 0
      wwd = 0
      ss = 0
      ffr = 0
      ee = 0
      rr(j) = 0
      rrs = 0
      rri = 0
      rrg = 0
      qsz = 0
      qiz = 0
      qgz = 0

   For i = 1 To lp
   
      mp1 = mp(i)
      pe = pp(i) - ek
      w = wp(i)
      wu = wup(i)
      wl = wlp(i)
      wd = wdp(i)
     
      Call yield(w, wu, wl, wd, pe, ek, e, eu, el, ed, r, ped, pedf, rd, nd, _
                  wm, wum, wlm, wdm, c, b, wmm, div, irs, iia, inn, ifc, imf, ef, ib)
      ep(i) = e
      eup(i) = eu
      elp(i) = el
      edp(i) = ed
      wp(i) = w
      wup(i) = wu
      wlp(i) = wl
      wdp(i) = wd
      rp(i) = r
      p = p + fp(i) * pp(i)
      ww = ww + fp(i) * wp(i)
      wwu = wwu + fp(i) * wup(i)
      wwl = wwl + fp(i) * wlp(i)
      wwd = wwd + fp(i) * wdp(i)
      ee = ee + fp(i) * ep(i)
      rr(j) = rr(j) + fp(i) * rp(i)

      fr = frp(i)
      s = sp(i)
      qqs = qsp(i)
      qqi = qip(i)
      qqg = qgp(i)
      ct = cp(i)
      Call divi3(fr, s, ct, pe, rd, ped, pedf, nd, qqg, qqi, qqs, rs, ri, rg, _
                 im, ki, kg, ex, sm, smm, cg, ci, irs, iia, inn, ifc, imf, ef, ib)
      frp(i) = fr
      sp(i) = s
      ffr = ffr + fp(i) * frp(i)
      ss = ss + fp(i) * sp(i)
      
      rsp(i) = rs
      rip(i) = ri
      rgp(i) = rg
      qsp(i) = qqs
      qip(i) = qqi
      qgp(i) = qqg
      rrs = rrs + fp(i) * rsp(i)
      rri = rri + fp(i) * rip(i)
      rrg = rrg + fp(i) * rgp(i)
      qss = qsp(i)
      qii = qip(i)
      qgg = qgp(i)
      rq = qss
      For jj = 1 To mp1 + 1
        qx(jj) = qxs(i, jj)
      Next jj
      Call musk(mp(i), rq, qx, c0, c1, c2)
      qss = rq
      For jj = 1 To mp1 + 1
        qxs(i, jj) = qx(jj)
      Next jj
      rq = qii
      For jj = 1 To mp1 + 1
        qx(jj) = qxi(i, jj)
      Next jj
      Call musk(mp(i), rq, qx, c0, c1, c2)
      qii = rq
      For jj = 1 To mp1 + 1
        qxi(i, jj) = qx(jj)
      Next jj
      rq = qgg
      For jj = 1 To mp1 + 1
         qx(jj) = qxg(i, jj)
      Next jj
      Call musk(mp(i), rq, qx, c0, c1, c2)
      qgg = rq
      For jj = 1 To mp1 + 1
        qxg(i, jj) = qx(jj)
      Next jj
      qsz = qsz + qss
      qiz = qiz + qii
      qgz = qgz + qgg
    
     it = sdsj(j)
     rdd.AddNew
     rdd(0) = NoYear
     rdd(1) = it
     rdd(2) = i
     rdd(3) = Int(ee * 100) / 100
     rdd(4) = Int(100 * pp(i)) / 100
     rdd(5) = Int(100 * wp(i)) / 100
     rdd(6) = Int(100 * wup(i)) / 100
     rdd(7) = Int(100 * wlp(i)) / 100
     rdd(8) = Int(100 * wdp(i)) / 100
     rdd(9) = Int(100 * qcal(j - 1) / lp) / 100
     rdd(10) = Int(100 * sp(i)) / 100
     rdd(11) = Int(100 * frp(i)) / 100
     rdd.Update
  
   Next i
      py = py + p
      emy = emy + em
      eky = eky + ek
      eey = eey + ee
      rsy = rsy + rrs
      riy = riy + rri
      rgy = rgy + rrg
      rrcc = rrcc + rr(j)
      pj(j) = p
      wj(j) = ww
      qsig(j + lag) = qsig(j + lag - 1) * cs + (qsz + qiz + qgz) * (1# - cs)
'      Write #2, CountFlood, rrcc
      qcal(j) = qsig(j)
    
    it = sdsj(j)
    rdd.AddNew
    rdd(0) = NoYear
    rdd(1) = it
    rdd(2) = 0
    rdd(3) = Int(ee * 100) / 100
    rdd(4) = Int(p * 100) / 100
    rdd(5) = Int(wp(1) * 100) / 100
    rdd(6) = Int(wwu * 100) / 100
    rdd(7) = Int(wwl * 100) / 100
    rdd(8) = Int(wwd * 100) / 100
    rdd(9) = Int(100 * qcal(j)) / 100
    rdd(10) = Int(100 * ss) / 100
    rdd(11) = Int(ffr * 100) / 100
    rdd.Update
Next j
rdd.Close
'Close #2

End Sub
