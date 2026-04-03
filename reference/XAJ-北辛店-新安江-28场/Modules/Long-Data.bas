Attribute VB_Name = "Long_Data"
Sub save_Data(m As Integer)
Dim i As Integer

bname = "Long_day"
b.CursorLocation = adUseClient
b.Open bname, cn, adOpenDynamic, adLockOptimistic

j = 1
Do While j <= m
        b.AddNew
        b(0) = glchsdsj(j)
        b.Update
       j = j + 1
Loop
b.Close

End Sub

