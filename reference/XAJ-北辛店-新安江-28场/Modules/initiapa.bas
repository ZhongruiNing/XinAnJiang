Attribute VB_Name = "initiapa"
Sub inialcheck()
Dim i As Integer, j As Integer, it As Long, nn As Integer
Dim m As Integer, sql1 As String

Dim iii As Integer, itb As Long, ite As Long
Dim iy As Integer, im As Integer, id As Integer, ih As Integer
Dim yy1 As Integer, mm1 As Integer, dd1 As Integer, hh1 As Integer
Dim yb As Integer, mb As Integer, db As Integer
Dim yy As Integer, mm As Integer, dd As Integer
Dim year As Integer
Dim PPa As Single

it = glchsdsj(1)
Call ymdh(it, iy, im, id, ih)
Call tymd(iy, im, id, -1, yy1, mm1, dd1)
Call yrsfd(yy1, mm1, dd1, ite)

bname = "dastlspa"
sql1 = "select * from  " + bname + " where dt =" + CStr(ite)
b.CursorLocation = adUseClient
b.Open sql1, cn, adOpenStatic

If b.EOF And b.BOF Then
    itb = 0
    iii = 1
Else
    iii = 0
    itb = b("dt")
    PPa = b(2)
    Exit Sub
End If
b.Close

If iii = 1 Then


bname = "dastls" + "pa"
sql1 = "select * from  " + bname + " where dt <= " + CStr(ite) + " order by dt desc"

b.CursorLocation = adUseClient
b.Open sql1, cn, adOpenStatic

If b.EOF And b.BOF Then
    itb = 0
    iii = 1
Else
    iii = 0
    itb = b("dt")
    PPa = b(2)
End If
b.Close

Call ymd(itb, year, mb, db)

If iii = 0 Then
  m = DateSerial(yy1, mm1, dd1) - DateSerial(year, mb, db)
  If m > BeginDay Then
     iii = 1
     itb = 0
  Else
    yy1 = year
    mm1 = mb
    dd1 = db
  End If
End If

If iii = 1 Then
  Call ymdh(it, iy, im, id, ih)
  Call tymd(iy, im, id, -BeginDay, yy1, mm1, dd1)
  If yy1 < iy Then
      yy1 = iy
      mm1 = 1
      dd1 = 1
   End If
End If
    
  Call ymdh(it, iy, im, id, ih)

  m = DateSerial(iy, im, id) - DateSerial(yy1, mm1, dd1)
  ReDim gsdsj(m), sdsj(m)

  yy = yy1
  mm = mm1
  dd = dd1
  Call yrsfd(yy, mm, dd, it)
  sdsj(0) = it
  gsdsj(0) = it
  For i = 1 To m
  Call tymd(yy, mm, dd, 1, yy1, mm1, dd1)
  Call yrsfd(yy1, mm1, dd1, it)
  sdsj(i) = it
  gsdsj(i) = it
  yy = yy1
  mm = mm1
  dd = dd1
  Next i
  
  Call calinial(iii, m, PPa)

Else
End If
End Sub
Sub calinial(iii As Integer, m As Integer, PPa As Single)
Dim nna As Integer
Dim p() As Single, pj() As Single, pa() As Single
Dim i As Integer, j As Integer, it As Long, kk As Integer, jj As Integer
Dim it1 As Long, it2 As Long, sql1 As String

Dim im As Single, Ka As Single

Dim cn  As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim b As New ADODB.Recordset

bname = "charactls"
rst.CursorLocation = adUseClient
rst.Open bname, cn, adOpenStatic
nna = rst.RecordCount
      im = rst(2)
      Ka = rst(3)
rst.Close

Na = nna

ReDim p(m, Na), pa(m), pj(m)
  
Call calcu_averagp_day(m)
Call inputrainPa(p, m)

For i = 1 To m
  pj(i) = 0
  For j = 1 To Na
   pj(i) = pj(i) + p(i, j) / Na
  Next j
Next i

If iii = 1 Then
   PPa = im / 3
Else
End If
pa(0) = PPa

For j = 1 To m
    kk = Ka
    pa(j) = Ka * (pa(j - 1) + pj(j))
   If pa(j) >= im Then
     pa(j) = im
   End If
  Next j


bname = "dastls" + "pa"
sql1 = "delete * from " + bname + " where dt between " + CStr(sdsj(2)) + " and  " + CStr(sdsj(m))
rst.CursorLocation = adUseClient
rst.Open bname, cn, adOpenDynamic, adLockOptimistic
cn.Execute (sql1)

For j = 1 To m
        rst.AddNew
        rst(0) = sdsj(j)
        rst(1) = pj(j)
        rst(2) = Int(pa(j) * 10) / 10
        rst.Update
Next j
rst.Close
End Sub

