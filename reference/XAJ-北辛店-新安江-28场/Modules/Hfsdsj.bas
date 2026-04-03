Attribute VB_Name = "hfsdsj"
Option Explicit
Function isleapyear(iy As Integer) As Boolean
'iy=yyyy年是否为闰年
If iy Mod 400 = 0 Or (iy Mod 100 <> 0 And iy Mod 4 = 0) Then
    isleapyear = True
Else
    isleapyear = False
End If
End Function
Function idy(iy As Integer, im As Integer) As Integer
'由iy年im月求该月天数iyd
If im = 4 Or im = 6 Or im = 9 Or im = 11 Then
idy = 30
Else
idy = 31
End If
If im = 2 Then
    If isleapyear(iy) Then
    idy = 29
    Else
    idy = 28
    End If
End If
End Function
Function idyyear(iy As Integer)
'由iy年求该年天数idyyear
If isleapyear(iy) Then
idyyear = 366
Else
idyyear = 365
End If
End Function

Function idyth(iy As Integer, im As Integer, id As Integer) As Integer
'由iy年im月id日ih时求距元旦idyth天
idyth = DateSerial(iy, im, id) - DateSerial(iy, 1, 1) + 1
End Function
Sub ymdh(it As Long, iy As Integer, im As Integer, id As Integer, ih As Integer)
'由it求iy年im月id日ih时
iy = it \ 1000000
im = ((it \ 10000) * 10000 - iy * 1000000) \ 10000
id = ((it \ 100) * 100 - CLng(iy) * 1000000 - CLng(im) * 10000) \ 100
ih = it - CLng(iy) * 1000000 - CLng(im) * 10000 - CLng(id) * 100
End Sub
Sub ymdhf(it As Double, iy As Integer, im As Integer, _
          id As Integer, ih As Integer, ff As Integer)
'Call ymdhf(it, iy, im, id, ih, ff)
'由it求iy年im月id日ih时ff分。it的格式为 1997070108.03 ,表示1997年7月1日8时03分。
Dim it1 As Long
iy = it \ 1000000
im = ((it \ 10000) * 10000 - iy * 1000000) \ 10000
id = ((it \ 100) * 100 - CLng(iy) * 1000000 - CLng(im) * 10000) \ 100
ih = it - CLng(iy) * 1000000 - CLng(im) * 10000 - CLng(id) * 100
it1 = Int(it)
ff = (it - it1) * 100
End Sub
Sub yrsf(iy As Integer, im As Integer, id As Integer, ih As Integer, it As Long)
'由iy年im月id日ih时求it
'it = iy * 1000000 + im * 10000 + id * 100 + ih
it = CLng(iy) * 1000000 + CLng(im) * 10000 + CLng(id) * 100 + CLng(ih)
End Sub
Sub yrsfd(iy As Integer, im As Integer, id As Integer, itd As Long)
'由iy年im月id日求it
'itd = iy * 10000 + im * 100 + id
itd = CLng(iy) * 10000 + CLng(im) * 100 + CLng(id)
End Sub
Sub ymd(it As Long, iy As Integer, im As Integer, id As Integer)
'由it求iy年im月id日
iy = it \ 10000
im = (it - CLng(iy) * 10000) \ 100
id = (it - CLng(iy) * 10000 - CLng(im) * 100)
End Sub
Sub tymdit(it As Long, idd As Integer, it1 As Long)
Dim iy As Integer, im As Integer, id As Integer
Dim iy1 As Integer, im1 As Integer, id1 As Integer
Call ymd(it, iy, im, id)
Call tymd(iy, im, id, idd, iy1, im1, id1)
Call yrsfd(iy1, im1, id1, it1)
End Sub
Sub tymd(iy As Integer, im As Integer, id As Integer, idd As Integer, iy1 As Integer, im1 As Integer, id1 As Integer)
'由iy年im月id日及增加的天数idd求iy1年im1月id1日
'-120<idd<120
Dim idt As Integer, imd As Integer
If idd <= 0 Then
    idt = id + idd
    If idt > 0 Then
    iy1 = iy
    im1 = im
    id1 = idt
    Else '1m
        If im = 1 Then
        im1 = 12
        iy1 = iy - 1
        Else
        im1 = im - 1
        iy1 = iy
        End If
        id1 = idy(iy1, im1) + idt
        If id1 <= 0 Then '2m
            If im1 = 1 Then
            im1 = 12
            iy1 = iy1 - 1
            Else
            im1 = im1 - 1
            iy1 = iy1
            End If
            id1 = idy(iy1, im1) + id1
        End If
        If id1 <= 0 Then '3m
            If im1 = 1 Then
            im1 = 12
            iy1 = iy1 - 1
            Else
            im1 = im1 - 1
            iy1 = iy1
            End If
            id1 = idy(iy1, im1) + id1
        End If
        If id1 <= 0 Then '4m
            If im1 = 1 Then
            im1 = 12
            iy1 = iy1 - 1
            Else
            im1 = im1 - 1
            iy1 = iy1
            End If
            id1 = idy(iy1, im1) + id1
        End If
    End If
Else 'idd>0
    idt = id + idd
    imd = idy(iy, im)
    If idt <= imd Then
        iy1 = iy
        im1 = im
        id1 = idt
    Else '
                        '1m
        If im < 12 Then
        iy1 = iy
        im1 = im + 1
        id1 = idt - imd
        Else 'im=12
        iy1 = iy + 1
        im1 = 1
        id1 = idt - imd
        End If
        imd = idy(iy1, im1)
        If id1 > imd Then '2m
           If im1 < 12 Then
            iy1 = iy1
            im1 = im1 + 1
            id1 = id1 - imd
            Else
            iy1 = iy1 + 1
            im1 = 1
            id1 = id1 - imd
            End If
        End If
        imd = idy(iy1, im1)
        If id1 > imd Then '3m
            If im1 < 12 Then
            iy1 = iy1
            im1 = im1 + 1
            id1 = id1 - imd
            Else
            iy1 = iy1 + 1
            im1 = 1
            id1 = id1 - imd
            End If
         End If
        imd = idy(iy1, im1)
        If id1 > imd Then '4m
            If im1 < 12 Then
            iy1 = iy1
            im1 = im1 + 1
            id1 = id1 - imd
            Else
            iy1 = iy1 + 1
            im1 = 1
            id1 = id1 - imd
            End If
        End If
    End If '
End If
End Sub
Sub tymdhit(it1 As Long, ihh As Integer, it2 As Long)
'由 it1 增加的时数ihh求 it2
Dim iy As Integer, im As Integer, id As Integer, ih As Integer
Dim iy1 As Integer, im1 As Integer, id1 As Integer, ih1 As Integer, ity As Long
Call ymdh(it1, iy, im, id, ih)
Call tymdh(iy, im, id, ih, ihh, iy1, im1, id1, ih1)
'Call yrsf(iy1, im1, id1, ih1, it2, ity)
End Sub
Sub tymdh(iy As Integer, im As Integer, id As Integer, ih As Integer, ihh As Integer, iy1 As Integer, im1 As Integer, id1 As Integer, ih1 As Integer)
'由iy年im月id日ih时及增加的时数ihh求iy1年im1月id1日ih1时
'-120*24<ihh<120*24        只有 24 时 无 0 时
Dim ii As Integer
ih1 = ih + ihh
If ih1 >= 0 Then
    If ih1 > 24 Then
    ii = ih1 \ 24
    ih1 = ih1 - ii * 24
    Call tymd(iy, im, id, ii, iy1, im1, id1)
    Else
    iy1 = iy
    im1 = im
    id1 = id
    End If
Else
    ii = -ih1 \ 24 + 1
    ih1 = ii * 24 + ih1
    Call tymd(iy, im, id, -ii, iy1, im1, id1)
End If
End Sub
Sub tymdh00(iy As Integer, im As Integer, id As Integer, ih As Integer, ihh As Integer, iy1 As Integer, im1 As Integer, id1 As Integer, ih1 As Integer)
'由iy年im月id日ih时及增加的时数ihh求iy1年im1月id1日ih1时
'-120*24<ihh<120*24        只有 0 时 无24 时
Dim ii As Integer
ih1 = ih + ihh
If ih1 >= 0 Then
    If ih1 >= 24 Then
    ii = ih1 \ 24
    ih1 = ih1 - ii * 24
    Call tymd(iy, im, id, ii, iy1, im1, id1)
    Else
    iy1 = iy
    im1 = im
    id1 = id
    End If
Else
    ii = -ih1 \ 24 + 1
    ih1 = ii * 24 + ih1
    Call tymd(iy, im, id, -ii, iy1, im1, id1)
End If
End Sub
Function nd12(iy1 As Integer, im1 As Integer, id1 As Integer, _
  ih1 As Integer, dt As Integer, iy2 As Integer, im2 As Integer, _
  id2 As Integer, ih2 As Integer) As Integer
  '由iy2年im2月id2日ih2时,时段长dt,iy1年im1月id1日ih1时， _
   求两时间间隔间的时段总数
Dim d12 As Integer, d12h As Integer
'Dim iy2 As Integer, im2 As Integer, id2 As Integer, ih2 As Integer
'Dim iy1 As Integer, im1 As Integer, id1 As Integer, ih1 As Integer
d12 = DateSerial(iy2, im2, id2) - DateSerial(iy1, im1, id1)
d12h = d12 * 24 + ih2 - ih1
nd12 = d12h \ dt
End Function

Function fd12h(iy1 As Integer, im1 As Integer, id1 As Integer, _
  ih1 As Integer, iy2 As Integer, im2 As Integer, _
  id2 As Integer, ih2 As Integer) As Integer
Dim d12 As Integer
d12 = DateSerial(iy2, im2, id2) - DateSerial(iy1, im1, id1)
fd12h = d12 * 24 + ih2 - ih1
End Function

Function fihe(ih As Integer, tt As Integer) As Integer
If ih >= 8 Then
fihe = 8 + ((ih - 8) \ tt) * tt
Exit Function
End If
If ih < 8 Then
fihe = ((ih + 16) \ tt) * tt - 16
   If fihe <= 0 Then
   fihe = 24 + fihe
   End If
End If
End Function

Sub subfihe(iy As Integer, im As Integer, id As Integer, ih As Integer, _
tt As Integer, iy1 As Integer, im1 As Integer, id1 As Integer, ih1 As Integer)
ih1 = fihe(ih, tt)
iy1 = iy: im1 = im: id1 = id
If ih1 > ih Then
Call tymd(iy, im, id, -1, iy1, im1, id1)
End If
End Sub
Sub zfcsj(symdh As String, it As Long)
Dim sy As String, sm As String, sd As String, sh As String, s As String
Dim iy As Integer, im As Integer, id As Integer, ih As Integer
Dim i1 As Integer
s = Mid(symdh, 1, 4)
i1 = Val(s)
iy = i1
s = Mid(symdh, 6, 2)
im = Val(s)
s = Mid(symdh, 9, 2)
id = Val(s)
s = Mid(symdh, 12, 2)
ih = Val(s)
it = CLng(iy) * 1000000 + CLng(im) * 10000 + CLng(id) * 100 + CLng(ih)
End Sub
 Sub sjzfc(iy As Integer, im As Integer, _
           id As Integer, ih As Integer, symdh As String)
 Dim sy As String, sm As String, sd As String, sh As String, s As String
 sy = Trim(Str(iy))
 If im < 10 Then
    sm = "0" + Trim(Str(im))
 Else
    sm = Trim(Str(im))
 End If
 If id < 10 Then
    sd = "0" + Trim(Str(id))
 Else
    sd = Trim(Str(id))
 End If
 If ih < 10 Then
    sh = "0" + Trim(Str(ih))
 Else
    sh = Trim(Str(ih))
 End If
 symdh = sy + "\" + sm + "\" + sd + Space(1) + sh
 symdh = Trim(symdh)
 End Sub
Sub zfcsj0(symdh As String, it As Long)
Dim sy As String, sm As String, sd As String, sh As String, s As String
Dim iy As Integer, im As Integer, id As Integer, ih As Integer
Dim i1 As Integer
s = Mid(symdh, 1, 2)
i1 = Val(s)
If i1 > 50 Then
    iy = i1 + 1900
Else
    iy = i1 + 2000
End If
s = Mid(symdh, 4, 2)
im = Val(s)
s = Mid(symdh, 7, 2)
id = Val(s)
s = Mid(symdh, 10, 2)
ih = Val(s)
it = CLng(iy) * 1000000 + CLng(im) * 10000 + CLng(id) * 100 + CLng(ih)
End Sub
 Sub sjzfc0(iy As Integer, im As Integer, _
           id As Integer, ih As Integer, symdh As String)
 Dim sy As String, sm As String, sd As String, sh As String, s As String
 s = Trim(Str(iy))
 sy = Right(s, 2)
 If im < 10 Then
    sm = "0" + Trim(Str(im))
 Else
    sm = Trim(Str(im))
 End If
 If id < 10 Then
    sd = "0" + Trim(Str(id))
 Else
    sd = Trim(Str(id))
 End If
 If ih < 10 Then
    sh = "0" + Trim(Str(ih))
 Else
    sh = Trim(Str(ih))
 End If
 symdh = sy + "\" + sm + "\" + sd + Space(1) + sh
 symdh = Trim(symdh)
 End Sub

Sub sjstrzh(it As Long, st As String)
Dim iy1 As Integer, im1 As Integer, id1 As Integer, ih1 As Integer
 Call ymdh(it, iy1, im1, id1, ih1)
 st = Trim(Str(im1)) + "月" + Trim(Str(id1)) + "日"
 st = Trim(st)
End Sub
Sub sjstrzh1(it As Long, st As String)
Dim iy1 As Integer, im1 As Integer, id1 As Integer, ih1 As Integer
 Call ymdh(it, iy1, im1, id1, ih1)
  If ih1 >= 10 Then
 st = Trim(Str(im1)) + "-" + Trim(Str(id1)) + " " + Trim(Str(ih1)) + ":00"
 Else
 st = Trim(Str(im1)) + "-" + Trim(Str(id1)) + " 0" + Trim(Str(ih1)) + ":00"
 End If
 st = Trim(st)
End Sub
