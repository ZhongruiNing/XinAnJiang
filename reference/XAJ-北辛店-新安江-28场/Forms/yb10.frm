VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form yb10 
   Caption         =   "蚌埠洪水预报计算结果"
   ClientHeight    =   6435
   ClientLeft      =   810
   ClientTop       =   1080
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   9615
   StartUpPosition =   2  '屏幕中心
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   5535
      Left            =   5280
      TabIndex        =   1
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   9763
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   9763
      _Version        =   393216
   End
   Begin VB.Menu mnuprint 
      Caption         =   "打印"
      Begin VB.Menu mnuybgc 
         Caption         =   "洪水预报过程线"
      End
      Begin VB.Menu mnuybtj 
         Caption         =   "洪水预报统计值"
      End
   End
End
Attribute VB_Name = "yb10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Dim dxx, dyy
If Screen.Width >= 12000 Then
dxx = Screen.Width * 0.01 * 0.9
Else
dxx = Screen.Width * 0.01 * 0.8
End If
If Screen.Height >= 9000 Then
dyy = Screen.Height * 0.01 * 1.1
Else
dyy = Screen.Height * 0.01 * 1.1
End If
'With MSFlexGrid1
'    .Top = 5 * dyy
'    .Height = 66 * dyy
'    '.Left = 15 * dxx
'    '.Width = 93 * dxx
'End With
'With MSFlexGrid2
'    .Top = 5 * dyy
'    .Height = 66 * dyy
'    '.Left = 15 * dxx
'    '.Width = 93 * dxx
'End With
Call Fd
End Sub

Sub Fd()
Dim itd As Long, sitd As String, i As Integer, glite As Integer
glite = 0
With MSFlexGrid1
.Rows = glnn + 1
.Cols = 6
.FixedCols = 0
.FixedRows = 1
.TextMatrix(0, 0) = "时段": .ColWidth(0) = 500:: .ColAlignment(0) = 4
.TextMatrix(0, 1) = "时    间": .ColWidth(1) = 1200: .ColAlignment(1) = 4
.TextMatrix(0, 2) = "面雨量": .ColWidth(2) = 600: .ColAlignment(2) = 4
.TextMatrix(0, 3) = "实测流量": .ColWidth(3) = 800: .ColAlignment(3) = 4
.TextMatrix(0, 4) = "预报流量": .ColWidth(4) = 800: .ColAlignment(4) = 4
.TextMatrix(0, 5) = "校正流量": .ColWidth(5) = 800: .ColAlignment(5) = 4
For i = 1 To glnn
.TextMatrix(i, 0) = Str(i)
 itd = glchsdsj(i)
 Call sjstrzh(itd, sitd)
.TextMatrix(i, 1) = sitd
If i <= glnn1 Then
.TextMatrix(i, 2) = Format(glpwhf(i), "##0.0")
.TextMatrix(i, 3) = Format(glqobshf(i), "##0.0")
End If
.TextMatrix(i, 4) = Format(glqcalhf(i), "##0.0")
.TextMatrix(i, 5) = Format(glqadjhf(i), "##0.0")
Next i
End With
'
Dim itd1 As Long, itd2 As Long, tt As Single
Dim qt1 As Single, qt2 As Single, qw1 As Single, qw2 As Single
tt = gltt
Call ybstatis(glqcalhf, glnn, itd1, qt1, itd2, qt2, tt, glnn1, qw1, qw2)

With MSFlexGrid2
.Rows = 11
.Cols = 4
.FixedCols = 0
.FixedRows = 1
.TextMatrix(0, 0) = "序号": .ColWidth(0) = 500: .ColAlignment(0) = 4
.TextMatrix(0, 1) = "统计项目": .ColWidth(1) = 1400: .ColAlignment(1) = 4
.TextMatrix(0, 2) = "内    容": .ColWidth(2) = 1200: .ColAlignment(2) = 4
.TextMatrix(0, 3) = "单    位": .ColWidth(3) = 1000: .ColAlignment(3) = 4
'
itd = glchsdsj(1)
Call sjstrzh(itd, sitd)
.TextMatrix(1, 0) = 1: .TextMatrix(1, 1) = "实测资料开始": .TextMatrix(1, 2) = sitd
itd = glchsdsj(glnn1)
Call sjstrzh(itd, sitd)
.TextMatrix(2, 0) = 2: .TextMatrix(2, 1) = "实测资料截止": .TextMatrix(2, 2) = sitd
itd = glchsdsj(glnn)
Call sjstrzh(itd, sitd)
.TextMatrix(3, 0) = 3: .TextMatrix(3, 1) = "预报计算截止": .TextMatrix(3, 2) = sitd
itd = itd1
Call sjstrzh(itd, sitd)
.TextMatrix(4, 0) = 4: .TextMatrix(4, 1) = "起涨时间": .TextMatrix(4, 2) = sitd
'
qt1 = 0.1 * Int(qt1 * 10)
.TextMatrix(5, 0) = 5: .TextMatrix(5, 1) = "起涨流量": .TextMatrix(5, 2) = Format(qt1, "##0.0")
.TextMatrix(5, 3) = " 立方米/秒"
'
itd = itd2
Call sjstrzh(itd, sitd)
.TextMatrix(6, 0) = 6: .TextMatrix(6, 1) = "预报峰现时间": .TextMatrix(6, 2) = sitd
'
qt2 = 0.1 * Int(qt2 * 10)
.TextMatrix(7, 0) = 7: .TextMatrix(7, 1) = "预报洪峰流量": .TextMatrix(7, 2) = Format(qt2, "##0.0")
.TextMatrix(7, 3) = " 立方米/秒"

.TextMatrix(8, 0) = 8: .TextMatrix(8, 1) = "预报洪水总量": .TextMatrix(8, 2) = Format(wcto, "##0.0000")
.TextMatrix(8, 3) = " 亿立方米"

.TextMatrix(9, 0) = 9: .TextMatrix(9, 1) = "实测洪水总量": .TextMatrix(9, 2) = Format(woto, "##0.0000")
.TextMatrix(9, 3) = " 亿立方米"

qw2 = (woto - wcto) / (Abs(woto) + 0.00001)
.TextMatrix(10, 0) = 10: .TextMatrix(10, 1) = "洪水总量误差": .TextMatrix(10, 2) = Format(qw2, "##0.0000")
.TextMatrix(10, 3) = " %"
If itd2 < glite Then
.Rows = 16
'
Dim robsy As Single, rcaly As Single, ce As Single, qom As Single, _
            qcm As Single, eqm As Single, iom As Integer, icm As Integer, _
            iem As Integer, dc As Single
Call chstatis(glqcalhf, glqobshf, glnn1, robsy, rcaly, ce, qom, _
                                 qcm, eqm, iom, icm, iem, dc, tt)
''output variables are 实测洪量 robsy,计算洪量 rcaly(万方),相对误差 ce; _
                       实测峰值 qom, 计算峰值 qcm, 相对误差 eqm; _
                       实测峰时 iom, 计算峰时icm,峰现时间预报误差 iem; _
                       确定性系数 dc
itd = glchsdsj(iom)
Call sjstrzh(itd, sitd)
.TextMatrix(6, 0) = 6: .TextMatrix(6, 1) = "实测峰现时间": .TextMatrix(6, 2) = sitd
.TextMatrix(6, 3) = ""
itd = glchsdsj(icm)
Call sjstrzh(itd, sitd)
.TextMatrix(7, 0) = 7: .TextMatrix(7, 1) = "预报峰现时间": .TextMatrix(7, 2) = sitd
.TextMatrix(7, 3) = ""
.TextMatrix(8, 0) = 8: .TextMatrix(8, 1) = "峰现时间误差": .TextMatrix(8, 2) = Str(iem)
.TextMatrix(8, 3) = "小时"
'
 qom = 0.1 * Int(qom * 10)
.TextMatrix(9, 0) = 9: .TextMatrix(9, 1) = "实测洪峰流量": .TextMatrix(9, 2) = Format(qom, "##0.0")
.TextMatrix(9, 3) = " 立方米/秒"
qcm = 0.1 * Int(qcm * 10)
.TextMatrix(10, 0) = 10: .TextMatrix(10, 1) = "预报洪峰流量": .TextMatrix(10, 2) = Format(qcm, "##0.0")
.TextMatrix(10, 3) = " 立方米/秒"
eqm = 0.01 * Int(eqm * 100)
.TextMatrix(11, 0) = 11: .TextMatrix(11, 1) = "洪峰相对误差": .TextMatrix(11, 2) = Format(eqm, "##0.00")
.TextMatrix(11, 3) = " %"
'
 robsy = Int(robsy * 10000) / 100000000
.TextMatrix(12, 0) = 12: .TextMatrix(12, 1) = "实测洪量": .TextMatrix(12, 2) = Format(robsy, "##0.0000")
.TextMatrix(12, 3) = " 亿立方米"
 rcaly = Int(rcaly * 10000) / 100000000
.TextMatrix(13, 0) = 13: .TextMatrix(13, 1) = "预报洪量": .TextMatrix(13, 2) = Format(rcaly, "##0.0000")
.TextMatrix(13, 3) = " 亿立方米"
 ce = 0.01 * Int(ce * 100)
.TextMatrix(14, 0) = 14: .TextMatrix(14, 1) = "洪量相对误差": .TextMatrix(14, 2) = Format(ce, "##0.00")
.TextMatrix(14, 3) = " %"
'
 qw1 = Int(qw1 * 10000) / 100000000
.TextMatrix(15, 0) = 15: .TextMatrix(15, 1) = "预报洪水总量": .TextMatrix(15, 2) = Format(qw1, "##0.0000")
.TextMatrix(15, 3) = " 亿立方米"
'
End If
End With
End Sub
Private Sub Form_Resize()
'Call tscal
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload yb10
yb09.Show
'chybht1.Refresh
End Sub
Sub dybg1()
Dim objTitle As String, gdAlignment() As Integer, _
      printpage As Boolean, ncols As Integer, j As Integer
printpage = True
 objTitle = dylyc + "洪水预报"
 ncols = yb10.MSFlexGrid1.Cols
 ReDim gdAlignment(ncols)
 For j = 1 To ncols
    gdAlignment(j) = 4
 Next j
Call PrintGridsub(yb10.MSFlexGrid1, objTitle, gdAlignment, _
     Printer, printpage)
'gdalignment(j),j=1 to gdcols——对齐方式，=1 左齐，4 居中，7 右齐
End Sub
Private Sub dybg2()
Dim objTitle As String, gdAlignment() As Integer, _
      printpage As Boolean, ncols As Integer, j As Integer
printpage = True
 objTitle = dylyc + "预报成果统计"
 ncols = yb10.MSFlexGrid2.Cols
 ReDim gdAlignment(ncols)
 For j = 1 To ncols
    gdAlignment(j) = 4
 Next j
Call PrintGridsub(yb10.MSFlexGrid2, objTitle, gdAlignment, _
     Printer, printpage)
     '
End Sub
Private Sub mnuybgc_Click()
Call dybg1
End Sub

Private Sub mnuybtj_Click()
Call dybg2
End Sub
