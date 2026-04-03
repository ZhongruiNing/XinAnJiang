VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form yb11 
   BackColor       =   &H00C0C0C0&
   Caption         =   "流量过程线分析图"
   ClientHeight    =   5445
   ClientLeft      =   810
   ClientTop       =   1080
   ClientWidth     =   8130
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   8130
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbarMain 
      Height          =   372
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2292
      _ExtentX        =   4048
      _ExtentY        =   661
      ButtonWidth     =   635
      ButtonHeight    =   582
      ImageList       =   "imgsTbar"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "select1"
            Object.ToolTipText     =   "查询预报依据和预报结果"
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "info1"
            Object.ToolTipText     =   "查询P、Q-T坐标"
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "grab1"
            Object.ToolTipText     =   "图形显示漫游"
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "zin"
            Object.ToolTipText     =   "时间坐标轴拉长"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "zout"
            Object.ToolTipText     =   "时间坐标轴收缩"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "print1"
            Object.ToolTipText     =   "结果打印"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      MousePointer    =   99
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   720
      ScaleHeight     =   6735
      ScaleWidth      =   10575
      TabIndex        =   0
      Top             =   960
      Width           =   10575
   End
   Begin ComctlLib.ImageList imgsTbar 
      Left            =   0
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "yb11.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "yb11.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "yb11.frx":0224
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "yb11.frx":0336
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "yb11.frx":0650
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "yb11.frx":096A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "yb11.frx":0A7C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnupopup 
      Caption         =   "popup"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnupop 
         Caption         =   "打印预报流量过程线图"
         Index           =   0
      End
      Begin VB.Menu mnupop 
         Caption         =   "打印雨量流量过程线"
         Index           =   1
      End
      Begin VB.Menu mnupop 
         Caption         =   "打印预报特征统计"
         Index           =   2
      End
   End
End
Attribute VB_Name = "yb11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As Database, rd1 As Recordset
Dim tx() As Single, pw() As Single, _
    qcal() As Single, qobs() As Single, qin() As Single, qadj() As Single
Dim ib As Integer, ie As Integer, nn As Integer, nnie As Integer
Dim ibe As Integer, iihh As Integer, numhh As Integer
Dim yhh() As Long, tnote() As String, nn1 As Integer
Dim itbar As Integer, iunload As Integer
Dim mdw As Integer, mup As Integer, x0 As Single, x1 As Single
Dim it0 As Long, lhh As Long
Dim phfxmin As Single, phfxmax As Single, phfymin As Single, phfymax As Single, _
     qhfymin As Single, qhfymax As Single, dxx As Single, pdyy As Single, qdyy As Single
Dim mecap As String
Sub picscal()
Dim dxxx, dyyy
dxxx = Me.Width * 0.01
dyyy = Me.Height * 0.01
Pic1.Top = dyyy * 3
Pic1.Height = dyyy * 90
Pic1.Left = dxxx * 5
Pic1.Width = dxxx * 90
End Sub
Sub subhtsj()
'输入绘图数据
Dim i As Integer, j As Integer, it As Long, iihh As Integer, ddnn As Integer
Dim iiy1 As Integer, iim1 As Integer, iid1 As Integer, iih1 As Integer
Dim iiy2 As Integer, iim2 As Integer, iid2 As Integer, iih2 As Integer, ity As Long
Dim sql1$, msg$
'On Error GoTo c
Me.Caption = dylyc1 + "流量过程线分析"
mecap = dylyc1 + "流量过程线分析"
ib = 1
ie = glnn
nn = glnn
nn1 = glnn1
ReDim tx(nn), pw(nn), qin(nn), _
      qcal(nn) As Single, qobs(nn), qin(nn), tnote(nn)
'
For i = 0 To nn
sdsj(i) = glchsdsj(i)
tx(i) = i
pw(i) = glpwhf(i)
qcal(i) = glqcalhf(i)
qobs(i) = glqobshf(i)
qadj(i) = glqcalhf(i)
Next i

it0 = sdsj(0)
qobs(0) = qobs(1)
qcal(0) = qcal(1)
qadj(0) = qadj(1)
iihh = -gltt
it = sdsj(1): Call ymdh(it, iiy1, iim1, iid1, iih1)
Call tymdh(iiy1, iim1, iid1, iih1, iihh, iiy2, iim2, iid2, iih2)
'Call yrsf(iiy2, iim2, iid2, iih2, it, ity)
sdsj(0) = it
tx(0) = 0
c: Exit Sub
End Sub

Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim it As Long, i As Integer, ii As Integer
Dim iiy As Integer, iim As Integer, iid As Integer, iih As Integer, ix As Integer
Dim yq As Long, scalzh As Single, yp As Long
'Dim myhourclass As Chourcross
'On Error GoTo c
If itbar = 3 Then
Me.Pic1.MousePointer = 99
Me.Pic1.MouseIcon = LoadPicture(Path + "\bmp\H_move.cur")
x0 = x
x1 = x
mdw = 1
End If
If itbar = 2 Then
            Me.Pic1.MousePointer = 99
            Me.Pic1.MouseIcon = LoadPicture(Path + "\bmp\cross_l.cur")
    If x > phfxmin And x < phfxmax Then
        ix = Int(x)
        If x - ix >= 0.5 Then
         ii = ix + 1
        Else
         ii = ix
        End If
it = sdsj(ii)
Call ymd(it, iiy, iim, iid)
mdh$ = Str(iim) & "月" & Str(iid) & "日"
'lbltx.Caption = mdh$
yq = Int(y)
scalzh = (phfymax - phfymin) / (40 * qdyy)
'lblqy.Caption = Str(yq) + "立方米/秒"
If y > qhfymin And y < qhfymax Then
MDImain!StatusBar1.Panels(1).Text = mdh$ + Space(4) + Str(yq) + "立方米/秒"
Me.Caption = mdh$ + Space(4) + Str(yq) + "立方米/秒"
End If
   End If
End If
c: Exit Sub
End Sub


Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim dx As Single, x2 As Single
ibe = ie - ib
If itbar = 3 And mdw = 1 Then
        x2 = Abs(x - x1)
        If x2 > (ie - ib) / 15# Then
        x1 = x
dx = x0 - x
If dx < 0# Then
   ib = ib + dx
   If ib < 1 Then
   ib = 1
   End If
   ie = ib + ibe
Else
   ie = ie + dx
   If ie > nn Then
   ie = nn
   End If
   ib = ie - ibe
End If

Call chyprocess(tx, pw, qobs, qcal, qin, qadj, nn, ib, ie, gltt, Me, phfxmin, phfxmax, phfymin, phfymax, _
               qhfymin, qhfymax, dxx, qdyy)
        End If
End If
End Sub
Private Sub Pic1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If itbar = 3 Then
mdw = 0
End If
If itbar = 2 Then
'lbltx.Caption = ""
'lblqy.Caption = ""
MDImain!StatusBar1.Panels(1).Text = ""
Me.Caption = mecap
End If
End Sub
Private Sub Form_Load()
Dim imgx As ListImage
Dim l As Integer, ynl As Long
itbar = 0
Call subhtsj
Set imgx = imgsTbar.ListImages.Add(1, , LoadPicture(Path + "\bmp\select.bmp"))
Set imgx = imgsTbar.ListImages.Add(2, , LoadPicture(Path + "\bmp\info2.bmp"))
Set imgx = imgsTbar.ListImages.Add(3, , LoadPicture(Path + "\bmp\grabber.bmp"))
Set imgx = imgsTbar.ListImages.Add(4, , LoadPicture(Path + "\bmp\zoomin.ico"))
Set imgx = imgsTbar.ListImages.Add(5, , LoadPicture(Path + "\bmp\zoomout.ico"))
'Set imgx = imgsTbar.ListImages.Add(6, , LoadPicture(Path1 + "\bmp\save1.bmp"))
Set imgx = imgsTbar.ListImages.Add(7, , LoadPicture(Path + "\bmp\print.bmp"))
iunload = 0
Me.AutoRedraw = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
iunload = 1
End Sub
Private Sub Form_Resize()
Call picscal
Call chyprocess(tx, pw, qobs, qcal, qin, qadj, nn, ib, ie, gltt, Me, phfxmin, phfxmax, phfymin, phfymax, _
               qhfymin, qhfymax, dxx, qdyy)
End Sub
Private Sub tbarMain_ButtonClick(ByVal Button As ComctlLib.Button)

'Public Sub tbarMain_ButtonClick(ByVal Button As Button)
Dim retval, optival
Dim ttlval As String, msgtxt As String
'Dim myhourclass As Chourcross
    Select Case Button.Key
        Case "select1"
            'yb04.Show vbModal
            itbar = 1
            tbarMain.Buttons.Item(1).Value = tbrUnpressed
            yb12.Show
        Case "info1"
            itbar = 2
            'Set myhourclass = New Chourcross
            Me.Pic1.MousePointer = 99
            Me.Pic1.MouseIcon = LoadPicture(Path + "\bmp\cross_l.cur")
        Case "grab1"
            itbar = 3
             Me.Pic1.MousePointer = 99
             Me.Pic1.MouseIcon = LoadPicture(Path + "\bmp\h_move.cur")
        Case "zin"
            itbar = 4
            tbarMain.Buttons.Item(4).Value = tbrUnpressed
            ib = ib + 24 \ gltt
            ie = ie - 24 \ gltt
            Call chyprocess(tx, pw, qobs, qcal, qin, qadj, nn, ib, ie, gltt, Me, phfxmin, phfxmax, phfymin, phfymax, _
               qhfymin, qhfymax, dxx, qdyy)
        Case "zout"
            itbar = 5
            tbarMain.Buttons.Item(5).Value = tbrUnpressed
            ib = ib - 24 \ gltt
            ie = ie + 24 \ gltt
           Call chyprocess(tx, pw, qobs, qcal, qin, qadj, nn, ib, ie, gltt, Me, phfxmin, phfxmax, phfymin, phfymax, _
               qhfymin, qhfymax, dxx, qdyy)
        'Case "save1"
        '    itbar = 6
        '    msgtxt = "预报结果是否发送到服务器入库？"
        '    ttlval = "确认预报结果入库"
        '    optival = vbExclamation + vbYesNo + vbDefaultButton2
        '    retval = MsgBox(msgtxt, optival, ttlval)
        '    If retval = vbYes Then
        '        'mdimain!StatusBar1.Panels(1).Text = "正在将预报结果入库......"
        '        Call jgrksub
        '        'mdimain!StatusBar1.Panels(1).Text = "预报结果已入库!"
        '    End If
        Case "print1"
            Call chybhthfprint(tx, pw, qobs, qcal, qin, nn, ib, ie, sdsj, gltt, phfxmin, phfxmax, phfymin, phfymax, _
               qhfymin, qhfymax, dxx, qdyy)
            Printer.EndDoc
            'PopupMenu mnupopup, 0
    End Select
    If itbar <> 2 And itbar <> 3 Then
    tbarMain.Buttons.Item(2).Value = tbrUnpressed
    tbarMain.Buttons.Item(3).Value = tbrUnpressed
    End If
End Sub

Private Sub mnupop_Click(Index As Integer)
'Dim prdxx, prdyy
    Select Case Index
        Case 0
            'Call dytx1
            'prdxx = Printer.ScaleWidth / 100#
            'prdyy = Printer.ScaleHeight / 100#
            'Printer.PaintPicture chybht1.Pic1, Printer.ScaleLeft + prdxx * 10, Printer.ScaleTop + prdyy * 10, prdxx * 80, prdyy * 50, _
            chybht1.Pic1.Left, chybht1.Pic1.Top, chybht1.Pic1.Width, chybht1.Pic1.Height, vbSrcPaint
            Call chybhthfprint(tx, pw, qobs, qcal, qin, nn, ib, ie, sdsj, gltt, phfxmin, phfxmax, phfymin, phfymax, _
               qhfymin, qhfymax, dxx, qdyy)
            Printer.EndDoc
        Case 1
             'Call dybg1
        Case 2
            'Call dybg2
    End Select
End Sub
