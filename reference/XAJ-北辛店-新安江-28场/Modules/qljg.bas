Attribute VB_Name = "qljg"
Sub qlybch()

bname = "ybresu" + dyly
b.CursorLocation = adUseClient
b.Open bname, cn
sql1 = "delete * from " + CStr(bname)
cn.Execute (sql1)

bname = "ybresu" + dyly + "1"
b.CursorLocation = adUseClient
b.Open bname, cn
sql1 = "delete * from " + CStr(bname)
cn.Execute (sql1)
b.Close

End Sub
Sub qlybri()

bname = "ybresu" + dyly
b.CursorLocation = adUseClient
b.Open bname, cn
sql1 = "delete * from " + CStr(bname)
cn.Execute (sql1)
b.Close

bname = "ybresu" + dyly + "1"
b.CursorLocation = adUseClient
b.Open bname, cn
sql1 = "delete * from  " + CStr(bname)
cn.Execute (sql1)
b.Close

End Sub

Sub qldast()

bname = "dast" + dyly
b.CursorLocation = adUseClient
b.Open bname, cn
sql1 = "delete * from " + CStr(bname)
cn.Execute (sql1)
b.Close

bname = "dast" + dyly
b.CursorLocation = adUseClient
b.Open bname, cn
sql1 = "delete * from " + CStr(bname)
cn.Execute (sql1)
b.Close

End Sub
