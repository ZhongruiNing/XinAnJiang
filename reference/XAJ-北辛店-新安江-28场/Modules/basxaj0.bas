Attribute VB_Name = "basxaj0"
Sub yield(w As Single, wu As Single, wl As Single, wd As Single, _
          pe As Single, ek As Single, e As Single, eu As Single, _
          el As Single, ed As Single, r As Single, ped() As Single, pedf() As Single, _
          rd() As Single, nd As Integer, wm As Single, wum As Single, _
          wlm As Single, wdm As Single, c As Single, b As Single, _
          wmm As Single, div As Single, _
          irs As Single, iia As Single, inn As Single, ifc As Single, imf As Single, _
          ef As Single, ib As Single)

Dim a As Single, peds As Single, rri As Single, fi As Single, ii As Integer
Dim minn As Single
 minn = 0.0001

      If pe <= div Then
         nd = 1
         ReDim ped(nd), pedf(nd)
         ped(1) = pe
      Else
         nd = Int(pe / div) + 1
         ReDim ped(nd), pedf(nd)
         For ii = 1 To nd - 1
          ped(ii) = div
         Next ii
          ped(nd) = pe - (nd - 1) * div + minn
       End If
       ReDim rd(nd)

    If (pe <= 0#) Then GoTo c1
      If w > wm Then w = wm
      a = wmm * (1# - (1# - w / wm) ^ (1# / (1# + b)))
      r = 0#
      peds = 0#
      For ii = 1 To nd
         a = a + ped(ii)
         peds = peds + ped(ii)
         rri = r
         r = peds - wm + w
         If a < wmm Then
         r = r + wm * (1# - a / wmm) ^ (1 + b)
         End If
         rd(ii) = r - rri
      Next ii
      fi = (iia * (wm - w) ^ inn + ifc) * (ef + 1)
      If (pe >= fi) Then
         irs = (pe - fi) * (1# - r / pe) * imf
      Else
         irs = (pe - fi / (ef + 1) * (1# - (1# - pe / fi) ^ (ef + 1))) * (1 - r / pe) * imf
      End If
      For ii = 1 To nd
         pedf(ii) = irs / nd * ib
      Next ii
      eu = ek
      el = 0
      ed = 0
      If (wu + pe - r) < wum Then GoTo c2
      If (wu + wl + pe - r - wum) >= wlm Then GoTo c3
      wl = wu + wl + pe - r - wum
      wu = wum
      GoTo c4
c3:   wu = wum
      wl = wlm
      wd = w + peds - r - wu - wl
      If wd > wdm Then wd = wdm
      GoTo c4
c2:   wu = wu + pe - r
      GoTo c4
c1:   r = 0
      If (wu + pe) < 0# Then GoTo c5
      eu = ek
      ed = 0
      el = 0
      wu = wu + pe
      GoTo c4
c5:   eu = wu + ek + pe
      wu = 0
      el = (ek - eu) * wl / wlm
      If wl < c * wlm Then el = c * (ek - eu)
      If (wl - el) < 0# Then GoTo c6
      ed = 0
      wl = wl - el
      GoTo c4
c6:   ed = el - wl
      el = wl
      wl = 0
      wd = wd - ed
c4:   w = wu + wl + wd
      e = eu + el + ed
      End Sub
Sub divi3(fr As Single, s As Single, ct As Single, _
          pe As Single, rd() As Single, ped() As Single, pedf() As Single, nd As Integer, _
          qqg As Single, qqi As Single, qqs As Single, rs As Single, _
          ri As Single, rg As Single, im As Single, ki As Single, kg As Single, _
          ex As Single, sm As Single, smm As Single, cg As Single, ci As Single, _
         irs As Single, iia As Single, inn As Single, ifc As Single, imf As Single, ef As Single, ib As Single)
     
     Dim rb As Single, rr As Single, kgd As Single, kid As Single, _
          td As Single, xx As Single, au As Single, ff As Single
     Dim ii As Integer
      If pe <= 0# Then GoTo c1
      rb = im * pe
      kid = (1# - (1# - (kg + ki)) ^ (1# / nd)) / (kg + ki)
      kgd = kid * kg
      kid = kid * ki
      rs = 0
      ri = 0
      rg = 0
      For ii = 1 To nd
        td = rd(ii) - im * ped(ii)
        xx = fr
        fr = td / ped(ii)
        If fr >= 1 Then
        fr = 1 - im
        End If
        s = xx * s / fr
        If (s >= sm) Then GoTo c2
        au = smm * (1# - (1# - s / sm) ^ (1# / (1# + ex)))
        ff = au + ped(ii) + pedf(ii) / fr * ib
      If ff < smm Then GoTo c3
c2:   rr = (ped(ii) + pedf(ii) / fr * ib + s - sm) * fr
      GoTo c4
c3:   ff = (1 - (ped(ii) + pedf(ii) / fr * ib + au) / smm) ^ (1 + ex)
      rr = (ped(ii) + pedf(ii) / fr * ib - sm + s + sm * ff) * fr
c4:   rs = rr + rs
      s = ped(ii) + pedf(ii) / fr * ib - rr / fr + s
      rg = s * kgd * fr + rg
      ri = s * kid * fr + ri
      s = s * (1# - kid - kgd)
      Next ii
      rs = rs + rb
      qqg = qqg * cg + rg * (1# - cg) * ct
      qqi = qqi * ci + ri * (1# - ci) * ct
      qqs = rs * ct
      Exit Sub
c1:   rs = 0#
      rg = s * kg * fr
      ri = s * ki * fr
      s = s * (1# - kg - ki)
      qqg = qqg * cg + rg * (1# - cg) * ct
      qqi = qqi * ci + ri * (1# - ci) * ct
      qqs = rs * ct
      End Sub
   Sub musk(MT As Integer, rq As Single, qx() As Single, _
               c0 As Single, c1 As Single, c2 As Single)
      
      Dim q1 As Single, q2 As Single, q3 As Single, jj As Integer, lm As Integer
      lm = MT + 1
      If lm = 1 Then GoTo c1
      For jj = 2 To lm
        q1 = rq
        q2 = qx(jj - 1)
        q3 = qx(jj)
        qx(jj - 1) = rq
        rq = c0 * q1 + c1 * q2 + c2 * q3
      Next jj
      qx(lm) = rq
      Exit Sub
c1:   qx(1) = rq
      End Sub
  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Charact_Watershed_Q(qcal() As Single, qobs() As Single, qin() As Single, m As Integer, _
            rin As Single, robsy As Single, rcaly As Single, ce As Single, qom As Single, _
            qcm As Single, eqm As Single, iom As Integer, icm As Integer, iem As Integer, _
            dc As Single, area As Single, dc2 As Single, dc3 As Single, dc4 As Single)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 'input variables are qcal(j), qobs(j),j=1 to m
 'output variables are 实测洪量 robsy,计算洪量 rcaly(万方),相对误差 ce _
                       实测峰值 qom, 计算峰值 qcm, 相对误差 eqm _
                       实测峰时 iom, 计算峰时icm,峰现时间预报误差 iem(小时) _
                       确定性系数 dc
 'input variables are tt,min1
 Dim rqo As Single, rqc As Single, eqobs As Single, f0 As Single, fn As Single
 Dim j As Integer, min1 As Single, tt As Single, rrc As Single
 
 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Dim f02 As Single, fn2 As Single, f03 As Single, fn3 As Single, rqoo As Single, eqobss As Single

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
   
       
     tt = gltt
     min1 = 0.000001
     
     ''''''''''''''''''''
     rqoo = 0#
     ''''''''''''''''''''
     
      rqo = 0#
      rqc = 0#
      rrc = 0
      For j = 1 To m
        rqo = rqo + qobs(j)
        rqc = rqc + qcal(j)
        
        ''''''''''''''''''''''''
        rqoo = rqoo + Sqr(qobs(j))
        ''''''''''''''''''''''''''''
        
        rrc = rrc + qin(j)
      Next j
      robsy = rqo * gltt * 3600 / area / 10 ^ 3
      rcaly = rqc * gltt * 3600 / area / 10 ^ 3

      ce = (robsy - rcaly) / (robsy + min1) * 100#
      eqobs = rqo / (m + min1)
      
      ''''''''''''''''''''''''''''''''
      eqobss = rqoo / (m + min1)
      ''''''''''''''''''''''''''''''''''''''''''''
      
      f0 = 0#
      fn = 0#
    ''''''''''''''''''''''''''''''
      f02 = 0#
      fn2 = 0#
      f03 = 0#
      fn3 = 0#
    '''''''''''''''''''''''''''
      
      For j = 1 To m
         f0 = f0 + (qobs(j) - eqobs) * (qobs(j) - eqobs)
         fn = fn + (qcal(j) - qobs(j)) * (qcal(j) - qobs(j))
         
         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         f02 = f02 + (Sqr(qobs(j)) - eqobss) * (Sqr(qobs(j)) - eqobss)
         fn2 = fn2 + (Sqr(qcal(j)) - Sqr(qobs(j))) * (Sqr(qcal(j)) - Sqr(qobs(j)))
         f03 = f03 + Abs(qobs(j) - eqobs)
         fn3 = fn3 + Abs(qcal(j) - qobs(j))
         '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         
      Next j
      f0 = f0 / (m + min1)
      fn = fn / (m + min1)
      
      '''''''''''''''''''''''''''''''''''''''''''
      f02 = f02 / (m + min1)
      fn2 = fn2 / (m + min1)
      f03 = f03 / (m + min1)
      fn3 = fn3 / (m + min1)
      '''''''''''''''''''''''''''''''''''''''''''
      
      dc = 1# - fn / (f0 + min1)
      
      ''''''''''''''''''''''''''''''''''''''''''
      dc2 = 1# - fn2 / (f02 + min1)
      dc3 = 1# - fn3 / (f03 + min1)
      dc4 = 1 - Abs(Sqr(rqc / rqo) - Sqr(rqo / rqc))
      '''''''''''''''''''''''''''''''''''''''''''''
      
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



Sub Charact_Watershed_W(qq() As Single, icm As Integer, ww() As Single)

 Dim i As Integer, j As Integer, tt As Single, jj As Integer, i1() As Integer, i2() As Integer, jp As Integer, ii() As Integer
 Dim it As Long, iy As Integer, im As Integer, id As Integer, ih As Integer
 Dim min As Single
 ReDim i1(5), i2(5), ww(5), ii(5)
 
 ii(1) = 1: ii(2) = 3: ii(3) = 5: ii(4) = 7
 
     jp = 5
     For i = 1 To jp
       i1(i) = 0
       i2(i) = 0
     Next i
     
  For j = 1 To 4
     
     i1(j) = icm - (24 / gltt) * ii(j) / 2
     If i1(j) <= 0 Then
       i1(j) = 1
     Else
     End If
     
     i2(j) = icm + (24 / gltt) * ii(j) / 2
     If i2(j) > LongTime Then
        i2(j) = LongTime
     Else
     End If
  
  Next j
     
    i1(5) = 1: i2(5) = LongTime
     
      For i = 1 To 5
            ww(i) = 0
         For j = i1(i) To i2(i)
           ww(i) = ww(i) + qq(j)
         Next j
     Next i
      
     For i = 1 To 5
        ww(i) = ww(i) * 0.36 * gltt / 10000 '(千方)
     Next i
     
         
   End Sub





Sub chPlotpw(tx() As Single, pw() As Single, nn As Integer, _
            ib As Integer, ie1 As Integer, ie2 As Integer, _
            plotpwcolor1 As Long, plotpwcolor2 As Long, chybht As Form, _
            phfymin As Single, phfymax As Single, qhfymin As Single, qhfymax As Single, qdyy As Single) '----------------
'chPlotpw(tx, pw, nn, ib, ie1, ie2, plotpwcolor1, plotpwcolor2)
'在 chybht!pic2 中画雨量过程线 pw(i),i=1,nn
Dim Xpoint As Single, Ypoint As Single
Dim x1 As Single, y1 As Single, ip As Integer
Dim i As Integer
chybht!Pic1.AutoRedraw = True
chybht!Pic1.DrawWidth = 1
chybht!Pic1.FillStyle = 0
chybht!Pic1.FillColor = plotpwcolor1
For i = ib To ie2
x1 = tx(i - 1)
y1 = qhfymax + 60 * qdyy
Xpoint = tx(i)
Ypoint = y1 - pw(i) * (60 * qdyy / (phfymax - phfymin))
    If pw(i) > 0# Then
If i <= ie1 Then
chybht!Pic1.Line (x1, y1)-(Xpoint, Ypoint), plotpwcolor1, BF
Else
chybht!Pic1.FillColor = plotpwcolor2
chybht!Pic1.Line (x1, y1)-(Xpoint, Ypoint), plotpwcolor2, BF
End If
    End If
Next i
End Sub

Sub Charact_Watershed_QD(qcal() As Single, qobs() As Single, qin() As Single, m As Integer, _
            rin As Single, robsy As Single, rcaly As Single, ce As Single, qom As Single, _
            qcm As Single, eqm As Single, iom As Integer, icm As Integer, _
            iem As Integer, dc As Single, area As Single)

 'input variables are qcal(j), qobs(j),j=1 to m
 'output variables are 实测洪量 robsy,计算洪量 rcaly(万方),相对误差 ce _
                       实测峰值 qom, 计算峰值 qcm, 相对误差 eqm _
                       实测峰时 iom, 计算峰时icm,峰现时间预报误差 iem(小时) _
                       确定性系数 dc
 'input variables are tt,min1
 Dim rqo As Single, rqc As Single, eqobs As Single, f0 As Single, fn As Single
 Dim j As Integer, min1 As Single, tt As Single, rrc As Single
   
       
     tt = gltt
     min1 = 0.000001
      rqo = 0#
      rqc = 0#
      rrc = 0
      For j = 1 To m
        rqo = rqo + qobs(j)
        rqc = rqc + qcal(j)
        rrc = rrc + qin(j)
      Next j
      'robsy = rqo * gltt * 3600 / 10 ^ 6
      'rcaly = rqc * gltt * 3600 / 10 ^ 6
      robsy = rqo * gltt * 3600 / area / 10 ^ 3
      rcaly = rqc * gltt * 3600 / area / 10 ^ 3

      ce = (robsy - rcaly) / (robsy + min1) * 100#
      eqobs = rqo / (m + min1)
      f0 = 0#
      fn = 0#
      For j = 1 To m
         f0 = f0 + (qobs(j) - eqobs) * (qobs(j) - eqobs)
         fn = fn + (qcal(j) - qobs(j)) * (qcal(j) - qobs(j))
      Next j
      f0 = f0 / (m + min1)
      fn = fn / (m + min1)
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




