Attribute VB_Name = "variety3"
Sub timeint()
Dim yy As Integer, mm As Integer, dd As Integer, hh As Integer, it As Long
Dim yy1 As Integer, mm1 As Integer, dd1 As Integer, hh1 As Integer
Dim i As Integer
ReDim glchsdsj(glnn)
it = StartTime
Call ymdh(it, yy, mm, dd, hh)
glchsdsj(0) = it
glchsdsj(1) = it
For i = 2 To glnn
Call tymdh(yy, mm, dd, hh, gltt, yy1, mm1, dd1, hh1)
Call yrsf(yy1, mm1, dd1, hh1, it)
glchsdsj(i) = it
yy = yy1
mm = mm1
dd = dd1
hh = hh1
Next i

ReDim glchml(glmm)
it = StartTime
Call ymdh(it, yy, mm, dd, hh)
glchml(0) = it
glchml(1) = it
For i = 2 To glmm
Call tymdh(yy, mm, dd, hh, gltm, yy1, mm1, dd1, hh1)
Call yrsf(yy1, mm1, dd1, hh1, it)
glchml(i) = it
yy = yy1
mm = mm1
dd = dd1
hh = hh1
Next i

Dim TimeAdd As Date
 
ReDim Hdate(glnn), glchmn(glnn)
Hdate(1) = TimeStart
glchmn(1) = TimeStart
For i = 2 To glnn
TimeAdd = DateAdd("h", gltt, Hdate(i - 1))
Hdate(i) = TimeAdd
glchmn(i) = TimeAdd
Next i


ReDim glchmm(glmm)

glchmm(0) = TimeStart
glchmm(1) = TimeStart
 For i = 2 To glmm
  HourAdd = DateAdd("h", gltm, glchmm(i - 1))
  glchmm(i) = HourAdd
Next i


End Sub
Sub timedayz()
Dim iyz As Integer, imz As Integer, idz As Integer, hh As Integer
Dim iym As Integer, imm As Integer, idm As Integer
Dim yy As Integer, mm As Integer, dd As Integer
Dim yy1 As Integer, mm1 As Integer, dd1 As Integer
Dim i As Integer, it As Long
Call ymdh(EndTime, iyz, imz, idz, hh)
Call ymdh(StartTime, iym, imm, idm, hh)
glday = DateSerial(iyz, imz, idz) - DateSerial(iym, imm, idm) + 1
glday = DateDiff("h", TimeStart, TimeEnd) / 24 + 2
ReDim sdsj(glday), gsdsj(glday)
Call ymdh(StartTime, yy, mm, dd, hh)
Call yrsfd(yy, mm, dd, it)
sdsj(0) = it
gsdsj(0) = it
sdsj(1) = it
gsdsj(1) = it
For i = 2 To glday
Call tymd(yy, mm, dd, 1, yy1, mm1, dd1)
Call yrsfd(yy1, mm1, dd1, it)
sdsj(i) = it
gsdsj(i) = it
yy = yy1
mm = mm1
dd = dd1
Next i
ReDim Ddata(glday)
Dim HourAdd As Date
Ddata(0) = TimeStart
Ddata(1) = TimeStart
For i = 2 To glday
  HourAdd = DateAdd("h", gltm, Ddata(i - 1))
  Ddata(i) = HourAdd
Next i
End Sub
Sub daytimed(m As Integer)
Dim iy As Integer, im As Integer, id As Integer
Dim yy1 As Integer, mm1 As Integer, dd1 As Integer
Dim i As Integer, it As Long
glnn = m
glnn1 = m
glday = m
ReDim gsdsj(m), glchsdsj(m), sdsj(m), glchmn(m)
it = StartTimeD
Call ymd(it, iy, im, id)
gsdsj(0) = it
gsdsj(1) = it
For i = 2 To LongTimeD
Call tymd(iy, im, id, 1, yy1, mm1, dd1)
Call yrsfd(yy1, mm1, dd1, it)
gsdsj(i) = it
iy = yy1
im = mm1
id = dd1
Next i
For i = 0 To LongTimeD
  glchsdsj(i) = gsdsj(i)
  sdsj(i) = gsdsj(i)
Next i

Dim HourAdd As Date
glchmn(1) = TimeStart
For i = 2 To m
  HourAdd = DateAdd("h", gltt, glchmn(i - 1))
  glchmn(i) = HourAdd
Next i


End Sub

