Attribute VB_Name = "watersurface"
Sub WatersurfaceD()

   Dim f As Single, tt As Single, i As Integer
   Dim zdylp() As Single, qcal() As Single, evap() As Single, m As Integer
   m = LongTimeD
   
   Dim pp() As Single

   ReDim evap(m), zdylp(m), qcal(m), pp(m)

   Call inputparawaterD(para1, para2, para3, Na)
   f = para1(3): tt = para1(2): Na = para1(4)
   
   Call eevapD(evap)

   Call inputrainwaterD(Na, zdylp, para3)


   Call waterD(evap, zdylp, qcal, pp, f, tt, m)

   Call savewaterD(evap, qcal, pp)
  
End Sub
Sub waterD(evap() As Single, zdylp() As Single, qcal() As Single, pp() As Single, f As Single, tt As Single, m As Integer)
    
Dim i As Integer, j As Integer, cp As Single
Dim pj As Single

cp = f / tt / 3.6
For i = 1 To m
  pp(i) = 0
  For j = 1 To Na
        pp(i) = pp(i) + zdylp(i, j) / Na
  Next j
Next i
      
For i = 1 To m
        pj = pp(i) - evap(i)
        If pj < 0 Then
          qcal(i) = 0
        Else
         qcal(i) = pj * cp
        End If
Next i

End Sub


