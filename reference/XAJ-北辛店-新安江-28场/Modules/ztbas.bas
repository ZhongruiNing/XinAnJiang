Attribute VB_Name = "ztbas"
Option Explicit
Sub cszcxp(itb As Long, iii As Integer, CountFlood As Integer)
Dim yyear As Integer, it1 As Long, it2 As Long
bname = "rain_day"
yyear = FloodNo(2, CountFlood)
Call yrsfd(yyear, 1, 1, it1)
Call yrsfd(yyear, 12, 31, it2)

sql1 = "select * from " + bname + " where dt between  " + CStr(it1) + " and " + CStr(it2) + "  order by DT  "
rd1.CursorLocation = adUseClient
rd1.Open sql1, cn

If rd1.EOF And rd1.BOF Then
itb = 0
iii = 1
Else
rd1.MoveFirst
iii = 0
itb = rd1("dt")
End If
rd1.Close
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
Sub cszcx1(itb As Long, itbb As Long)
Dim m_db As Database, rd1 As Recordset
Dim sql1 As String
Set m_db = OpenDatabase(Path + "\state\state.mdb")
sql1 = "select  *  from " + " dastly" + _
       "  where dt<" + CStr(itbb) + "  and dt>=" + CStr(itb) + "  order by dt  desc  "
Set rd1 = m_db.OpenRecordset(sql1, dbOpenDynaset)
If rd1.EOF And rd1.BOF Then
itb = 0
Else

itb = rd1(0)
End If
End Sub
Sub cszcx2(it1 As Long, ifail As Integer)
Dim m_db As Database, rd1 As Recordset
Dim sql1 As String
Set m_db = OpenDatabase(Path + "\state\state.mdb")
sql1 = "select  *  from " + " dastly" + "  where dt =" + CStr(it1)
Set rd1 = m_db.OpenRecordset(sql1, dbOpenDynaset)
If rd1.EOF And rd1.BOF Then
ifail = 1
Else
ifail = 0
End If
End Sub

