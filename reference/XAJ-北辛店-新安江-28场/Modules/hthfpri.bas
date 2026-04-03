Attribute VB_Name = "hthfpri"
Option Explicit

Sub chqaxisnumbers(Htics As Integer, Vtics As Integer, numbercolor As Long, _
                  ib As Integer, sdsj() As Long, nn As Integer, tt As Integer, chybht As Form, _
    phfxmin As Single, phfxmax As Single, qhfymin As Single, qhfymax As Single)
Dim xnuminc As Single, ynuminc As Single
Dim xnum1 As Single, ynum1 As Single
Dim gridspacex As Single, gridspacey As Single
Dim x1 As Single, y1 As Single
Dim ix As Integer, iy As Integer
Dim iiy As Integer, iim As Integer, iid As Integer, iih As Integer
Dim i As Integer, it As Long
Dim dy As Long, iidd As Integer
Dim mdh$, Number$, m1$, d1$, h1$
'iidd 为每个较大标记间包含的时段数
'tt 为次洪模型时段长
iidd = 24 / tt
chybht!Pic1.ForeColor = numbercolor
'xnuminc = chybht!Pic1.ScaleWidth / (Htics - 1)
ynuminc = (qhfymax - qhfymin) / (Vtics - 1)
'xnum1 = chybht!Pic1.ScaleLeft
ynum1 = qhfymin
If Htics <= 1 Then GoTo c
gridspacex = (phfxmax - phfxmin) / (Htics - 1)
gridspacey = Int((qhfymax - qhfymin) / (Vtics - 1) / 10) * 10 '(qhfymax - qhfymin) / (Vtics - 1)
ynuminc = gridspacey
For ix = 0 To Htics - 1 Step 3
    i = (ib - 1) + ix * iidd
    it = sdsj(i + 1)
    Call ymdh(it, iiy, iim, iid, iih)
    m1$ = CStr(iim)
    d1$ = CStr(iid)
    h1$ = CStr(iih)
    '
    Number$ = (d1$) + "日"
    x1 = phfxmin + ix * gridspacex - 0.5 * chybht.Pic1.TextWidth(Number$)
    y1 = qhfymin + 0.5 * chybht.Pic1.TextHeight("0")
    chybht.Pic1.CurrentX = x1
    chybht.Pic1.CurrentY = y1
    chybht.Pic1.Print Number$
    'mdh$ = Trim(m1$) + "月" + Trim(d1$) + "日"
    mdh$ = (m1$) + "月"
    x1 = phfxmin + ix * gridspacex - 0.5 * chybht.Pic1.TextWidth(mdh$)
    y1 = qhfymin + 0.5 * chybht.Pic1.TextHeight("0")
    chybht.Pic1.CurrentX = x1
    chybht.Pic1.CurrentY = y1 + 1 * chybht.Pic1.TextHeight("0")
    chybht.Pic1.Print mdh$
Next ix
c: For iy = 0 To Vtics - 1
    dy = Int(Abs(ynum1) + 0.00001) * sign(ynum1) + _
         iy * Int(Abs(ynuminc) + 0.00001) * sign(ynuminc)
    Number$ = (dy)
    x1 = phfxmin - chybht.Pic1.TextWidth(Number$) - chybht.Pic1.TextWidth(".")
    y1 = qhfymin + iy * gridspacey - 0.5 * chybht.Pic1.TextHeight("0")
    chybht.Pic1.CurrentX = x1
    chybht.Pic1.CurrentY = y1
    chybht.Pic1.Print Number$
Next
          
    x1 = phfxmin - chybht.Pic1.TextWidth(Number$) - chybht.Pic1.TextWidth(" 流量 ")
    y1 = qhfymin + (iy - 1) * gridspacey - 0.5 * chybht.Pic1.TextHeight("0")
    chybht.Pic1.CurrentX = x1
    chybht.Pic1.CurrentY = y1
    chybht.Pic1.Print "流量"

End Sub

Sub chqaxisnumbersD(Htics As Integer, Vtics As Integer, numbercolor As Long, _
                  ib As Integer, sdsj() As Long, nn As Integer, tt As Integer, chybht As Form, _
    phfxmin As Single, phfxmax As Single, qhfymin As Single, qhfymax As Single)
Dim xnuminc As Single, ynuminc As Single
Dim xnum1 As Single, ynum1 As Single
Dim gridspacex As Single, gridspacey As Single
Dim x1 As Single, y1 As Single
Dim ix As Integer, iy As Integer
Dim iiy As Integer, iim As Integer, iid As Integer, iih As Integer
Dim i As Integer, it As Long
Dim dy As Long, iidd As Integer
Dim mdh$, Number$, m1$, d1$, h1$
'iidd 为每个较大标记间包含的时段数
'tt 为次洪模型时段长
iidd = 24 / tt
chybht!Pic1.ForeColor = numbercolor
'xnuminc = chybht!Pic1.ScaleWidth / (Htics - 1)
ynuminc = (qhfymax - qhfymin) / (Vtics - 1)
'xnum1 = chybht!Pic1.ScaleLeft
ynum1 = qhfymin
If Htics <= 1 Then GoTo c
gridspacex = (phfxmax - phfxmin) / (Htics - 1)
gridspacey = (qhfymax - qhfymin) / (Vtics - 1)
For ix = 0 To Htics - 1 Step 20
    i = (ib - 1) + ix * iidd
    it = sdsj(i + 1)
    Call ymd(it, iiy, iim, iid)
    m1$ = CStr(iim)
    d1$ = CStr(iid)
    
    Number$ = (d1$) + "日"
    x1 = phfxmin + ix * gridspacex - 0.5 * chybht.Pic1.TextWidth(Number$)
    y1 = qhfymin + 0.5 * chybht.Pic1.TextHeight("0")
    chybht.Pic1.CurrentX = x1
    chybht.Pic1.CurrentY = y1
    chybht.Pic1.Print Number$
    'mdh$ = Trim(m1$) + "月" + Trim(d1$) + "日"
    mdh$ = (m1$) + "月"
    x1 = phfxmin + ix * gridspacex - 0.5 * chybht.Pic1.TextWidth(mdh$)
    y1 = qhfymin + 0.5 * chybht.Pic1.TextHeight("0")
    chybht.Pic1.CurrentX = x1
    chybht.Pic1.CurrentY = y1 + 1 * chybht.Pic1.TextHeight("0")
    chybht.Pic1.Print mdh$
Next ix
c: For iy = 0 To Vtics - 1
    dy = Int(Abs(ynum1) + 0.00001) * sign(ynum1) + _
         iy * Int(Abs(ynuminc) + 0.00001) * sign(ynuminc)
    Number$ = (dy)
    x1 = phfxmin - chybht.Pic1.TextWidth(Number$) - chybht.Pic1.TextWidth(".")
    y1 = qhfymin + iy * gridspacey - 0.5 * chybht.Pic1.TextHeight("0")
    chybht.Pic1.CurrentX = x1
    chybht.Pic1.CurrentY = y1
    chybht.Pic1.Print Number$
Next
          
    x1 = phfxmin - chybht.Pic1.TextWidth(Number$) - chybht.Pic1.TextWidth("  流量 ")
    y1 = qhfymin + (iy - 1) * gridspacey - 0.5 * chybht.Pic1.TextHeight("0")
    chybht.Pic1.CurrentX = x1
    chybht.Pic1.CurrentY = y1
    chybht.Pic1.Print "流量"

End Sub

Function sign(x As Single) As Integer
If x >= 0 Then
   sign = 1
Else
   sign = -1
End If
End Function

Sub chpaxisnumbersprint(Vtics As Integer, numbercolor As Long, _
     phfxmin As Single, phfxmax As Single, phfymin As Single, phfymax As Single, dxx As Single, _
     qhfymin As Single, qhfymax As Single, qdyy As Single)
Dim ynuminc As Single
Dim ynum1 As Single
Dim gridspacey As Single
Dim x1 As Single, y1 As Single
Dim ix As Integer, iy As Integer
Dim dy As Integer, Number$
'printer.ForeColor = numbercolor
ynuminc = (phfymax - phfymin) / (Vtics - 1)
ynum1 = phfymin
gridspacey = (60 * qdyy) / (Vtics - 1)
For iy = 0 To Vtics - 2
    dy = Int(Abs(ynum1) + 0.00001) * sign(ynum1) + _
         iy * Int(Abs(ynuminc) + 0.00001) * sign(ynuminc)
    Number$ = Format$(dy)
    x1 = phfxmin - Printer.TextWidth(Number$) - Printer.TextWidth(".")
    y1 = qhfymax + 60 * qdyy - iy * gridspacey - 0.5 * Printer.TextHeight("0")
    Printer.CurrentX = x1
    Printer.CurrentY = y1
    Printer.Print Number$
Next
End Sub

Sub chqaxisnumbersprint(Htics As Integer, Vtics As Integer, numbercolor As Long, _
                  ib As Integer, sdsj() As Long, nn As Integer, tt As Integer, _
    phfxmin As Single, phfxmax As Single, qhfymin As Single, qhfymax As Single)
Dim xnuminc As Single, ynuminc As Single
Dim xnum1 As Single, ynum1 As Single
Dim gridspacex As Single, gridspacey As Single
Dim x1 As Single, y1 As Single
Dim ix As Integer, iy As Integer
Dim iiy As Integer, iim As Integer, iid As Integer, iih As Integer
Dim i As Integer, it As Long
Dim dy As Long, iidd As Integer
Dim mdh$, Number$, m1$, d1$, h1$
'iidd 为每个较大标记间包含的时段数
'tt 为次洪模型时段长
iidd = 24 / tt
'printer.ForeColor = numbercolor
'xnuminc = printer.ScaleWidth / (Htics - 1)
ynuminc = (qhfymax - qhfymin) / (Vtics - 1)
'xnum1 = printer.ScaleLeft
ynum1 = qhfymin
If Htics <= 1 Then GoTo c
gridspacex = (phfxmax - phfxmin) / (Htics - 1)
gridspacey = (qhfymax - qhfymin) / (Vtics - 1)
For ix = 0 To Htics - 1
    i = (ib - 1) + ix * iidd
    it = sdsj(i)
    Call ymdh(it, iiy, iim, iid, iih)
    m1$ = CStr(iim)
    d1$ = CStr(iid)
    h1$ = CStr(iih)
    '
    Number$ = Trim(d1$) + "日"
    x1 = phfxmin + ix * gridspacex - 0.5 * Printer.TextWidth(Number$)
    y1 = qhfymin + 0.5 * Printer.TextHeight("0")
    Printer.CurrentX = x1
    Printer.CurrentY = y1
    Printer.Print Number$
    'mdh$ = Trim(m1$) + "月" + Trim(d1$) + "日"
    mdh$ = Trim(m1$) + "月"
    x1 = phfxmin + ix * gridspacex - 0.5 * Printer.TextWidth(mdh$)
    y1 = qhfymin + 0.5 * Printer.TextHeight("0")
    Printer.CurrentX = x1
    Printer.CurrentY = y1 + 1 * Printer.TextHeight("0")
    Printer.Print mdh$
Next ix
c: For iy = 0 To Vtics - 1
    dy = Int(Abs(ynum1) + 0.00001) * sign(ynum1) + _
         iy * Int(Abs(ynuminc) + 0.00001) * sign(ynuminc)
    Number$ = Format$(dy)
    x1 = phfxmin - Printer.TextWidth(Number$) - Printer.TextWidth(".")
    y1 = qhfymin + iy * gridspacey - 0.5 * Printer.TextHeight("0")
    Printer.CurrentX = x1
    Printer.CurrentY = y1
    Printer.Print Number$
Next
End Sub

Sub chpgridonprint(Vtics As Integer, minorticcolor As Long, _
     phfxmin As Single, phfxmax, phfymin As Single, phfymax As Single, dxx As Single, _
     qhfymin As Single, qhfymax As Single, qdyy As Single)
' Vtics - 1= 主要标记坐标区间数
Dim gridspacex As Single, gridspacey As Single
Dim minorspaceX As Single, minorspaceY As Single
Dim minorXtic As Single, minorYtic As Single, Xtic As Single
Dim xfrom As Single, Yfrom As Single, Xto As Single, Yto As Single
Dim iy As Integer, jy As Integer
Dim iiyy As Integer
iiyy = 5 '主要标记间小标记划分的区间数
'Printer.AutoRedraw = True
gridspacey = 60 * qdyy / (Vtics - 1)
minorspaceY = gridspacey / iiyy
minorXtic = Abs(phfxmax - phfxmin) / 100: Xtic = minorXtic * 2
Printer.DrawStyle = 0
'Printer.ForeColor = minorticcolor
For iy = 1 To Vtics - 1
    xfrom = phfxmin
    Xto = xfrom + Xtic
    Yfrom = qhfymax + qdyy * 60 - iy * gridspacey
    Yto = Yfrom
    Printer.Line (xfrom, Yfrom)-(Xto, Yto)
Next iy
For jy = 1 To iiyy * (Vtics - 1)
    If (jy \ iiyy) * iiyy <> jy Then
    xfrom = phfxmin
    Yfrom = qhfymax + qdyy * 60 - jy * minorspaceY
    Xto = xfrom + minorXtic
    Yto = Yfrom
    Printer.Line (xfrom, Yfrom)-(Xto, Yto)
    End If
Next jy
End Sub
Sub chqgridonprint(Htics As Integer, Vtics As Integer, minorticcolor As Long, _
     phfxmin As Single, phfxmax As Single, qhfymin As Single, qhfymax As Single)
' Htics - 1=,Vtics - 1= 主要标记坐标区间数
Dim gridspacex As Single, gridspacey As Single
Dim minorspaceX As Single, minorspaceY As Single
Dim minorXtic As Single, minorYtic As Single, Xtic As Single, Ytic As Single
Dim xfrom As Single, Yfrom As Single, Xto As Single, Yto As Single
Dim ix As Integer, iy As Integer, jx As Integer, jy As Integer
Dim iixx As Integer, iiyy As Integer
'iixx = 5
iixx = 8
iiyy = 5
'Printer.AutoRedraw = True
If Htics > 1 Then
gridspacex = (phfxmax - phfxmin) / (Htics - 1) 'X坐标主要标记间的距离
End If
gridspacey = (qhfymax - qhfymin) / (Vtics - 1)  'Y坐标主要标记间的距离
minorspaceX = gridspacex / iixx 'X坐标较小标记间的距离
minorspaceY = gridspacey / iiyy 'Y坐标较小标记间的距离
'minorXtic = Abs(printer.ScaleWidth) / 50
minorXtic = Abs(phfxmax - phfxmin) / 100  '较小标记的水平长度
minorYtic = Abs(qhfymax - qhfymin) / 50  '较小标记的垂直长度
Xtic = minorXtic * 2: Ytic = minorYtic * 2
Printer.DrawStyle = 0
'Printer.ForeColor = minorticcolor
For ix = 1 To Htics - 1
    xfrom = phfxmin + ix * gridspacex
    Xto = xfrom
    Yfrom = qhfymin
    Yto = Yfrom + Ytic
    Printer.Line (xfrom, Yfrom)-(Xto, Yto)
Next ix
For iy = 1 To Vtics - 1
    xfrom = phfxmin
    Xto = phfxmin + Xtic
    Yfrom = qhfymin + iy * gridspacey
    Yto = Yfrom
    Printer.Line (xfrom, Yfrom)-(Xto, Yto)
Next iy
For jx = 1 To iixx * (Htics - 1)
    If (jx \ iixx) * iixx <> jx Then
    xfrom = phfxmin + jx * minorspaceX
    Yfrom = qhfymin
    Xto = xfrom
    Yto = Yfrom + minorYtic
    Printer.Line (xfrom, Yfrom)-(Xto, Yto) '画X坐标较小标记
    End If
Next
For jy = 1 To iiyy * (Vtics - 1)
    If (jy \ iiyy) * iiyy <> jy Then
    xfrom = phfxmin
    Yfrom = qhfymin + jy * minorspaceY
    Xto = xfrom + minorXtic
    Yto = Yfrom
    Printer.Line (xfrom, Yfrom)-(Xto, Yto) '画Y坐标较小标记
    End If
Next
End Sub

Sub chPlotpwprint(tx() As Single, pw() As Single, nn As Integer, _
            ib As Integer, ie1 As Integer, ie2 As Integer, _
            plotpwcolor1 As Long, plotpwcolor2 As Long, _
            phfymin As Single, phfymax As Single, qhfymin As Single, qhfymax As Single, qdyy As Single) '----------------
'chPlotpw(tx, pw, nn, ib, ie1, ie2, plotpwcolor1, plotpwcolor2)
'在 chybht!pic2 中画雨量过程线 pw(i),i=1,nn
Dim Xpoint As Single, Ypoint As Single
Dim x1 As Single, y1 As Single, ip As Integer
Dim i As Integer
'Printer.AutoRedraw = True
Printer.DrawWidth = 2
Printer.FillStyle = 1
'Printer.FillColor = RGB(255, 255, 255)
For i = ib To ie2
x1 = tx(i - 1)
y1 = qhfymax + 60 * qdyy
Xpoint = tx(i)
Ypoint = y1 - pw(i) * (60 * qdyy / (phfymax - phfymin))
    If pw(i) > 0# Then
If i <= ie1 Then
Printer.Line (x1, y1)-(Xpoint, Ypoint), plotpwcolor1, B
Else
'Printer.FillColor = plotpwcolor2
Printer.Line (x1, y1)-(Xpoint, Ypoint), plotpwcolor1, B
End If
    End If
Next i
End Sub

Sub chPlotffqprint(tx() As Single, q() As Single, nn As Integer, _
           ib As Integer, ie1 As Integer, ie2 As Integer, icbsj As Integer, _
            plotqcolor1 As Long, plotqcolor2 As Long, idrawstyle As Integer, dxx As Single)  '-------------------
'在 printer 中画流量过程线 ffq(tx):q(i),i=1,nn
'chPlotffq(tx, q, nn , ib , ie1 , ie2 , icbsj ,plotpwcolor1 , plotqcolor2)
Dim Xpoint As Single, Ypoint As Single
Dim x1 As Single, y1 As Single, x2 As Single
Dim i As Integer, ib1 As Integer
'printer.AutoRedraw = True
Printer.DrawWidth = 2
If idrawstyle >= 0 Then
Printer.DrawStyle = idrawstyle
End If
ib1 = ib - icbsj
If ib1 <= 1 Then ib1 = 1
x1 = tx(ib1 - 1) + icbsj
x2 = tx(ib1 - 1)
y1 = ffq(x2, tx, q, nn)
Printer.CurrentX = x1
Printer.CurrentY = y1
Printer.ForeColor = plotqcolor1
For i = ib1 To ie2
Xpoint = tx(i) + icbsj
x2 = tx(i)
Ypoint = ffq(x2, tx, q, nn)
If i > ie1 Then
Printer.ForeColor = plotqcolor2
End If
    If Ypoint > 0# Then
Printer.Line -(Xpoint, Ypoint)
If idrawstyle = -1 Then
Printer.FillStyle = 1
Printer.Circle (Xpoint, Ypoint), dxx / 2
End If
     End If
Next i
End Sub


Sub chprinterscale(tx() As Single, q1() As Single, q2() As Single, q3() As Single, nn As Integer, _
               ib As Integer, ie As Integer, qhfymin As Single, qhfymax As Single, _
               dxx As Single, dyy As Single) '-------------------
Dim newpoint As Single, i As Integer
Dim imax As Integer
Dim Xmin0 As Single, Xmax0 As Single
Dim Ymin0 As Single, Ymax0 As Single
Xmin0 = tx(ib - 1): Xmax0 = tx(ie)
Ymin0 = qhfymin
Ymax0 = qhfymax
imax = 0
For i = 1 To nn
    newpoint = q1(i)
    If newpoint > Ymax0 Then
    Ymax0 = newpoint
    imax = 1
    End If
    If newpoint < Ymin0 And newpoint > 1# Then Ymin0 = newpoint
Next i
For i = 1 To nn
    newpoint = q2(i)
    If newpoint > Ymax0 Then
    Ymax0 = newpoint
    imax = 1
    End If
    If newpoint < Ymin0 And newpoint > 1# Then Ymin0 = newpoint
Next
For i = 1 To nn
    newpoint = q3(i)
    If newpoint > Ymax0 Then
    Ymax0 = newpoint
    imax = 1
    End If
    If newpoint < Ymin0 And newpoint > 1# Then Ymin0 = newpoint
Next
If imax = 1 Then
Ymax0 = Ymax0 + 10#
End If
Ymax0 = 10# * Int(Ymax0 / 10)
dxx = (Xmax0 - Xmin0) * 0.01
dyy = (Ymax0 - Ymin0) * 0.01
qhfymin = Ymin0
qhfymax = Ymax0
Xmin0 = Xmin0 - dxx * 15
Xmax0 = Xmax0 + dxx * 5
Ymax0 = Ymax0 + dyy * 120
Ymin0 = Ymin0 - dyy * 80
Printer.Scale (Xmin0, Ymax0)-(Xmax0, Ymin0)
'printer.AutoRedraw = True
Printer.DrawWidth = 1
'printer.ForeColor = RGB(0, 0, 0)
Printer.Line (Xmin0 + dxx * 15, qhfymin)-(Xmax0 - dxx * 5, qhfymin)
Printer.Line (Xmin0 + dxx * 15, qhfymin)-(Xmin0 + dxx * 15, qhfymax + dyy * 60)
Printer.Line (Xmax0 - dxx * 5, qhfymin)-(Xmax0 - dxx * 5, qhfymax + dyy * 60)
Printer.Line (Xmin0 + dxx * 15, qhfymax + dyy * 60)-(Xmax0 - dxx * 5, qhfymax + dyy * 60)

'Ymin1 = Ymin0: Ymax1 = Ymax0
End Sub

Sub chybhthfprint(tx() As Single, pw() As Single, qobs() As Single, _
            qcal() As Single, qadj() As Single, nn As Integer, _
            ib As Integer, ie As Integer, sdsj() As Long, tt As Integer, _
            phfxmin As Single, phfxmax As Single, phfymin As Single, phfymax As Single, _
            qhfymin As Single, qhfymax As Single, dxx As Single, qdyy As Single)
Dim plotcolor1 As Long, plotcolor2 As Long, minorticcolor As Long, numbercolor As Long
Dim Htics As Integer, Vtics As Integer, iidd As Integer
Dim ie1 As Integer, ie2 As Integer, icbsj As Integer, ibe As Integer
Dim idrawstyle As Integer
Dim txt As String, x1 As Single, y1 As Single, x2 As Single, y2 As Single
'On Error GoTo c
'tx(ib-1)=ib-1 tx(ie)=ie  tx(i) = i
'sdsj(ib-1) ,sdsj(ie) 为图形显示的点过程起止时间.
'格式最好为 sdsj(ib-1)=yyyymmdd08,sdsj(ie)=yyyymmdd08
' Htics - 1 ,Vtics - 1 为主要标记坐标区间数
' iidd 为每个较大标记间包含的时段数
'tt 为次洪模型时段长
ibe = ie - ib
If ibe >= nn Or ibe <= 1 Then Exit Sub
If nn < 1 Then Exit Sub
If ib < 1 Then
    ib = 1
    ie = ib + ibe
End If
If ie > nn Then
    ie = nn
    ib = ie - ibe
End If
'Printer.Cls
'xtitle = "时        间 (小 时)"
'ytitle = "      流 量 .立方米/秒."
'pytitle = "雨 量.毫米."
Htics = 1 + Int((ie - ib + 1) * tt / 24)
'ie = ib - 1 + (Htics - 1) * 24 \ tt
Vtics = 6
'iidd = int((ie - ib + 1) / (Htics - 1))
'iidd = 24 / tt
'dyy = (printer.Height) * 0.01
'dxx = (printer.Width) * 0.01
Screen.MousePointer = 11
dxx = (tx(ie) - tx(ib - 1)) * 0.01
phfxmin = tx(ib - 1)
phfxmax = tx(ie)
'
qhfymin = 0#
qhfymax = 20#
Call chprinterscale(tx, qcal, qobs, qobs, nn, ib, ie, qhfymin, qhfymax, dxx, qdyy)
plotcolor1 = RGB(255, 255, 0) '黄
plotcolor2 = RGB(255, 255, 0) '青
ie1 = LongTime: ie2 = ie: icbsj = 0
idrawstyle = 4 'vbdashdotdot
Call chPlotffqprint(tx, qadj, nn, ib, ie1, ie2, icbsj, plotcolor1, plotcolor2, idrawstyle, qdyy)
'chybht!Lbladj.ForeColor = RGB(255, 255, 0) '黄
'chybht!Lbladj.Caption = "----预报校正流量"
    txt = " 来水流量"
    x1 = phfxmin + dxx * 60
    y1 = (qhfymin - 20 * qdyy)
    x2 = x1 + dxx * 10
    y2 = y1
    Printer.ForeColor = plotcolor1
    Printer.DrawStyle = idrawstyle
    Printer.Line (x1, y1)-(x2, y2)
    Printer.CurrentX = x2
    Printer.CurrentY = y1 - 0.5 * Printer.TextHeight("0")
    Printer.Print txt
'
plotcolor1 = RGB(255, 0, 0) '红
plotcolor2 = RGB(255, 0, 0) '红
ie1 = LongTimeD: ie2 = ie: icbsj = 0
idrawstyle = 0 'vbsodid
Call chPlotffqprint(tx, qcal, nn, ib, ie1, ie2, icbsj, plotcolor1, plotcolor2, idrawstyle, qdyy)
'chybht!lblcal.ForeColor = RGB(255, 0, 0) '红
'chybht!lblcal.Caption = "----预报计算流量"
    txt = " 预报流量"
    x1 = phfxmin + dxx * 30
    y1 = (qhfymin - 20 * qdyy)
    x2 = x1 + dxx * 10
    y2 = y1
    Printer.ForeColor = plotcolor1
    Printer.DrawStyle = idrawstyle
    Printer.Line (x1, y1)-(x2, y2)
    Printer.CurrentX = x2
    Printer.CurrentY = y1 - 0.5 * Printer.TextHeight("0")
    Printer.Print txt
'
plotcolor1 = RGB(0, 0, 255) '蓝
plotcolor2 = RGB(0, 0, 255) '蓝
ie1 = LongTimeD: ie2 = ie1: icbsj = 0
If ie <= LongTimeD Then
    ie2 = ie
Else
    ie2 = LongTimeD
End If
idrawstyle = -1 'vbdash
Call chPlotffqprint(tx, qobs, nn, ib, ie1, ie2, icbsj, plotcolor1, plotcolor2, idrawstyle, dxx)
'chybht!lblobs.ForeColor = RGB(0, 0, 255) '蓝
'chybht!lblobs.Caption = "---- 实测入库流量"
    txt = " 实测流量"
    x1 = phfxmin
    y1 = (qhfymin - 20 * qdyy)
    x2 = x1 + dxx * 10
    y2 = y1
    Printer.ForeColor = plotcolor1
    Printer.Line (x1, y1)-(x2, y2)
    Printer.FillStyle = 1
    Printer.Circle (x1 + dxx * 5, y1), dxx / 2
    Printer.CurrentX = x2
    Printer.CurrentY = y1 - 0.5 * Printer.TextHeight("0")
    Printer.Print txt
'
minorticcolor = RGB(0, 0, 0)
Call chqgridonprint(Htics, Vtics, minorticcolor, phfxmin, phfxmax, qhfymin, qhfymax)
numbercolor = RGB(0, 0, 0)
Call chqaxisnumbersprint(Htics, Vtics, numbercolor, ib, sdsj, nn, tt, phfxmin, phfxmax, qhfymin, qhfymax)

phfymin = 0#
phfymax = 10#
Call chPic2scale(tx, pw, nn, ib, ie, phfymin, phfymax, dxx, qhfymin, qhfymax, qdyy, yb09)
plotcolor1 = RGB(0, 0, 255) '蓝
plotcolor2 = RGB(0, 255, 255) '青
ie1 = LongTimeD
If ie <= LongTimeD Then
    ie2 = ie
Else
    ie2 = LongTimeD
End If
Call chPlotpwprint(tx, pw, nn, ib, ie1, ie2, plotcolor1, plotcolor2, phfymin, phfymax, qhfymin, qhfymax, qdyy)
minorticcolor = RGB(0, 0, 0)
Call chpgridonprint(Vtics, minorticcolor, phfxmin, phfxmax, phfymin, phfymax, dxx, qhfymin, qhfymax, qdyy)
numbercolor = RGB(0, 0, 0)
Call chpaxisnumbersprint(Vtics, numbercolor, phfxmin, phfxmax, phfymin, phfymax, dxx, qhfymin, qhfymax, qdyy)
    
    txt = dylyc + "雨量流量合成图"
    x1 = phfxmin + 30 * dxx
    y1 = (qhfymax + 80 * qdyy)
    Printer.FontSize = 15
    Printer.ForeColor = RGB(0, 0, 0) '蓝
    Printer.CurrentX = x1
    Printer.CurrentY = y1 - 0.5 * Printer.TextHeight("0")
    Printer.Print txt
    x1 = phfxmin + 40 * dxx
    y1 = (qhfymax + 70 * qdyy)
    Printer.FontSize = 10
    Printer.ForeColor = RGB(0, 0, 0) '蓝
    Printer.CurrentX = x1
    Printer.CurrentY = y1 - 0.5 * Printer.TextHeight("0")
    Printer.FontSize = 10

Screen.MousePointer = 1
c: Exit Sub
End Sub
Sub chqaxisnumberspro(Htics As Integer, Vtics As Integer, numbercolor As Long, _
                  ib As Integer, sdsj() As Long, nn As Integer, tt As Integer, chybht As Form, _
    phfxmin As Single, phfxmax As Single, qhfymin As Single, qhfymax As Single)
Dim xnuminc As Single, ynuminc As Single
Dim xnum1 As Single, ynum1 As Single
Dim gridspacex As Single, gridspacey As Single
Dim x1 As Single, y1 As Single
Dim ix As Integer, iy As Integer
Dim iiy As Integer, iim As Integer, iid As Integer, iih As Integer
Dim i As Integer, it As Long
Dim dy As Long, iidd As Integer
Dim mdh$, Number$, m1$, d1$, h1$
'iidd 为每个较大标记间包含的时段数
'tt 为次洪模型时段长
iidd = 24 / tt
chybht!Pic1.ForeColor = numbercolor
'xnuminc = chybht!Pic1.ScaleWidth / (Htics - 1)
ynuminc = (1.1 * qhfymax - 1.1 * qhfymin) / (Vtics - 1)
'xnum1 = chybht!Pic1.ScaleLeft
ynum1 = qhfymin
If Htics <= 1 Then GoTo c
gridspacex = (phfxmax - phfxmin) / (Htics - 1)
gridspacey = 1.1 * (qhfymax - qhfymin) / (Vtics - 1)
For ix = 0 To Htics - 1 Step 1
    i = (ib - 1) + ix * iidd
    it = sdsj(i + 1)
    Call ymdh(it, iiy, iim, iid, iih)
    m1$ = CStr(iim)
    d1$ = CStr(iid)
    h1$ = CStr(iih)
    '
    Number$ = (d1$) + "日"
    x1 = phfxmin + ix * gridspacex - 0.5 * chybht.Pic1.TextWidth(Number$)
    y1 = qhfymin + 0.5 * chybht.Pic1.TextHeight("0")
    chybht.Pic1.CurrentX = x1
    chybht.Pic1.CurrentY = y1
    chybht.Pic1.Print Number$
    'mdh$ = Trim(m1$) + "月" + Trim(d1$) + "日"
    mdh$ = (m1$) + "月"
    x1 = phfxmin + ix * gridspacex - 0.5 * chybht.Pic1.TextWidth(mdh$)
    y1 = qhfymin + 0.5 * chybht.Pic1.TextHeight("0")
    chybht.Pic1.CurrentX = x1
    chybht.Pic1.CurrentY = y1 + 1 * chybht.Pic1.TextHeight("0")
    chybht.Pic1.Print mdh$
Next ix
c: For iy = 0 To Vtics - 1
    dy = Int(Abs(ynum1) + 0.00001) * sign(ynum1) + _
         iy * Int(Abs(ynuminc) + 0.00001) * sign(ynuminc)
    Number$ = (dy)
    x1 = phfxmin - chybht.Pic1.TextWidth(Number$) - chybht.Pic1.TextWidth(".")
    y1 = qhfymin + (iy + 1) * gridspacey + 2.5 * chybht.Pic1.TextHeight("0")
    chybht.Pic1.CurrentX = x1
    chybht.Pic1.CurrentY = y1
    chybht.Pic1.Print Number$
Next
End Sub

