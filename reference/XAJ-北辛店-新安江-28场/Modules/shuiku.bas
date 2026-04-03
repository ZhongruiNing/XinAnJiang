Attribute VB_Name = "shuiku"
Sub qinflowh(ia As Integer, qc() As Single, qrc() As Single)


Dim i As Integer, j As Integer, k As Integer, m As Integer, jj As Integer

Dim c0 As Single, c1 As Single, c2 As Single, tt As Single, q0() As Single, x0() As Single, n0() As Integer
Dim q() As Single, xx() As Single, nn() As Integer, kk() As Long
Dim qoxs() As Single

tt = gltt
m = glnn

ReDim qc(glnn), qrc(glnn)
ReDim xx(ia), nn(ia), kk(ia)
Dim inname() As String, qmax As Single, ii As Integer
ReDim q(ia, glnn), qoxs(glnn), qc(glnn), qrc(glnn), inname(ia)

If ia <> 0 Then

For j = 1 To m
    qrc(j) = 0
Next j
 
 bname = "Rout_muskingum_" + CStr(dyly)
 sql1 = "select * from " + bname + " order by [序号]  "
 rd1.CursorLocation = adUseClient
 rd1.Open sql1, cn
 jj = rd1.RecordCount
 ReDim q0(jj), x0(jj), n0(jj)
For k = 1 To ia
          
     If Not rd1.BOF Then
         jj = rd1.RecordCount
             q0(k) = rd1(1)
             x0(k) = rd1(2)
             n0(k) = rd1(3)
             rd1.MoveNext
     Else
         MsgBox dylyc + "河道演算参数文件不存在，请建立该文件！"
         rd1.Close
          Exit Sub
     End If
 
     inname(k) = para3(Na + k)
 
      bname = "qdata" + dyly
      it1 = glchsdsj(1)
      it2 = glchsdsj(glnn)
      sql1 = "select " + CStr(inname(k)) + " from " + bname + " where dt between  " + CStr(it1) + " and " + CStr(it2) + "  order by DT  "
      b.CursorLocation = adUseClient
      b.Open sql1, cn

      If Not b.BOF Then
          For j = 1 To glnn
             If b(0) >= 0.0001 Then
                q(k, j) = b(0)
             Else
                q(k, j) = 0
             End If
             b.MoveNext
          Next j
      End If
      b.Close
      
       
      
      For j = 1 To glnn
         qc(j) = 0#
      Next j

      qmax = q(k, 1)
      For j = 1 To glnn
         If qmax < q(k, j) Then
            qmax = q(k, j)
         Else
         End If
      Next j

      ii = 1
      For j = 1 To jj
         If qmax > q0(j) Then
            ii = j
         Else
         End If
      Next j

      xx(k) = x0(k): kk(k) = tt: nn(k) = n0(k)
      xx(k) = 0.5 - nn(k) * (0.5 - xx(k))

      c0 = (0.5 * tt - kk(k) * xx(k)) / (kk(k) - kk(k) * xx(k) + 0.5 * tt)
      c1 = (kk(k) * xx(k) + 0.5 * tt) / (kk(k) - kk(k) * xx(k) + 0.5 * tt)
      c2 = (kk(k) - kk(k) * xx(k) - 0.5 * tt) / (kk(k) - kk(k) * xx(k) + 0.5 * tt)
     
      For j = 1 To m
         qoxs(j) = q(k, j)
      Next j
      For j = 1 To m
         qc(j) = 0
      Next j
     
      qc(1) = Qjliu / ia / 3
     
   If nn(k) <> 0 Then

          For i = 1 To nn(k)
              For j = 2 To m
                qc(j) = c0 * qoxs(j) + c1 * qoxs(j - 1) + c2 * qc(j - 1)
              Next j
              If i = nn(k) Then GoTo ee
              For j = 1 To m
                qoxs(j) = qc(j)
              Next j
          Next i

   Else
          For j = 1 To m
            qc(j) = q(k, j)
          Next j
   End If

ee:   For i = 1 To m
          qrc(i) = qrc(i) + qc(i)
      Next i
                
  Next k
  rd1.Close

      For j = 1 To m
           qc(j) = 0#
           For k = 1 To ia
             qc(j) = qc(j) + q(k, j)
           Next k
       Next j

Else
End If
End Sub
Sub qinflowD(ia As Integer, qc() As Single, qrc() As Single)


Dim i As Integer, j As Integer, k As Integer, m As Integer, jj As Integer

Dim c0 As Single, c1 As Single, c2 As Single, tt As Single, q0() As Single, x0() As Single, n0() As Integer
Dim q() As Single, xx() As Single, nn() As Integer, kk() As Long
Dim qoxs() As Single

tt = gltt
m = glnn

ReDim qc(glnn), qrc(glnn)
ReDim xx(ia), nn(ia), kk(ia)
Dim inname() As String, qmax As Single, ii As Integer
ReDim q(ia, glnn), qoxs(glnn), qc(glnn), qrc(glnn), inname(ia)

If ia <> 0 Then

For j = 1 To m
    qrc(j) = 0
Next j
 
 bname = "Rout_muskingum_" + CStr(dyly)
 sql1 = "select * from " + bname + " order by [序号]  "
 rd1.CursorLocation = adUseClient
 rd1.Open sql1, cn
 jj = rd1.RecordCount
 ReDim q0(jj), x0(jj), n0(jj)
For k = 1 To ia
          
     If Not rd1.BOF Then
         jj = rd1.RecordCount
             q0(k) = rd1(1)
             x0(k) = rd1(2)
             n0(k) = rd1(3)
             rd1.MoveNext
     Else
         MsgBox dylyc + "河道演算参数文件不存在，请建立该文件！"
         rd1.Close
          Exit Sub
     End If
 
     inname(k) = para3(Na + k)
 
      bname = "qdata_day" + dyly
      it1 = glchsdsj(1)
      it2 = glchsdsj(glnn)
      sql1 = "select " + CStr(inname(k)) + " from " + bname + " where dt between  " + CStr(it1) + " and " + CStr(it2) + "  order by DT  "
      b.CursorLocation = adUseClient
      b.Open sql1, cn

      If Not b.BOF Then
          For j = 1 To glnn
             If b(0) >= 0.0001 Then
                q(k, j) = b(0)
             Else
                q(k, j) = 0
             End If
             b.MoveNext
          Next j
      End If
      b.Close
      
      
      
      
      For j = 1 To glnn
         qc(j) = 0#
      Next j

      qmax = q(k, 1)
      For j = 1 To glnn
         If qmax < q(k, j) Then
            qmax = q(k, j)
         Else
         End If
      Next j

      ii = 1
      For j = 1 To jj
         If qmax > q0(j) Then
            ii = j
         Else
         End If
      Next j

      xx(k) = x0(k): kk(k) = tt: nn(k) = n0(k)
      xx(k) = 0.5 - nn(k) * (0.5 - xx(k))

      c0 = (0.5 * tt - kk(k) * xx(k)) / (kk(k) - kk(k) * xx(k) + 0.5 * tt)
      c1 = (kk(k) * xx(k) + 0.5 * tt) / (kk(k) - kk(k) * xx(k) + 0.5 * tt)
      c2 = (kk(k) - kk(k) * xx(k) - 0.5 * tt) / (kk(k) - kk(k) * xx(k) + 0.5 * tt)
     
      For j = 1 To m
         qoxs(j) = q(k, j)
      Next j
      For j = 1 To m
         qc(j) = 0
      Next j
     
      qc(1) = Qjliu / ia / 3
     
   If nn(k) <> 0 Then

          For i = 1 To nn(k)
              For j = 2 To m
                qc(j) = c0 * qoxs(j) + c1 * qoxs(j - 1) + c2 * qc(j - 1)
              Next j
              If i = nn(k) Then GoTo ee
              For j = 1 To m
                qoxs(j) = qc(j)
              Next j
          Next i

   Else
          For j = 1 To m
            qc(j) = q(k, j)
          Next j
   End If

ee:   For i = 1 To m
          qrc(i) = qrc(i) + qc(i)
      Next i
                
  Next k
  rd1.Close

      For j = 1 To m
           qc(j) = 0#
           For k = 1 To ia
             qc(j) = qc(j) + q(k, j)
           Next k
       Next j

Else
End If
End Sub
