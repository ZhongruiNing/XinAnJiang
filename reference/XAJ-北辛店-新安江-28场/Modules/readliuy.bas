Attribute VB_Name = "readliuy"

Sub readjliu()
Dim i As Integer, j As Integer
Dim it1 As Long, it2 As Long

ReDim Q00(3)

bname = "qdata"
it1 = glchsdsj(1)
sql1 = "select [灵口],[卢氏] from " + bname + " where [dt]=" + CStr(glchsdsj(1))
b.CursorLocation = adUseClient
b.Open sql1, cn

If b.BOF Or b.EOF Then
  MsgBox "实测流量资料文件不存在，请先建立该文件！"
  Exit Sub
Else
  For i = 0 To 1
    Q00(i + 1) = b(i)
  Next i
End If
b.Close
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


