VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDImain 
   BackColor       =   &H80000003&
   Caption         =   "伊洛河流域水文模型率定系统"
   ClientHeight    =   6765
   ClientLeft      =   165
   ClientTop       =   705
   ClientWidth     =   9120
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6390
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   10425
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2009-8-31"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "09:47"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnudata 
      Caption         =   "模拟时段和流域选择"
   End
   Begin VB.Menu mnutime 
      Caption         =   "洪水选取"
   End
   Begin VB.Menu mnurimd 
      Caption         =   "日模型计算"
      Begin VB.Menu r0 
         Caption         =   "-"
      End
      Begin VB.Menu mnurihz 
         Caption         =   "日模型调试"
      End
      Begin VB.Menu r2 
         Caption         =   "-"
      End
      Begin VB.Menu mnurish 
         Caption         =   "结果显示"
      End
      Begin VB.Menu r3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuztxs 
         Caption         =   "土壤含水量流域状态"
      End
   End
   Begin VB.Menu mnuybwy 
      Caption         =   "洪水模型"
      Begin VB.Menu b0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuycx 
         Caption         =   "降雨-径流次洪模拟"
      End
      Begin VB.Menu b1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuybta 
         Caption         =   "模拟结果图形显示"
      End
      Begin VB.Menu b6 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuxtwh 
      Caption         =   "系统维护"
      Begin VB.Menu mnuwucd 
         Caption         =   "日模型参数"
      End
      Begin VB.Menu w0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuwuch 
         Caption         =   "次洪模型参数"
      End
      Begin VB.Menu hh 
         Caption         =   "-"
      End
      Begin VB.Menu mnust00 
         Caption         =   "流域初始状态"
      End
      Begin VB.Menu n6 
         Caption         =   "-"
      End
      Begin VB.Menu mnucscx 
         Caption         =   "清理流域状态"
      End
      Begin VB.Menu p0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuqlch 
         Caption         =   "清理次洪预报结果"
      End
      Begin VB.Menu n10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuqlri 
         Caption         =   "清理日预报结果"
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "退出"
   End
End
Attribute VB_Name = "MDImain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuchpp_Click()
yb02.Show
End Sub
Private Sub mnudass_Click()
yb06.Show
End Sub
Private Sub mnulsybjg_Click()
yb06.Show
End Sub

Private Sub mnuana2_Click()
 Call interp
End Sub

Private Sub mnucscx_Click()
  Call qldast
End Sub

Private Sub mnudata_Click()
yb08.Show
End Sub

Private Sub mnuexit_Click()
 End
 cn.Close
End Sub
Private Sub mnuhysh_Click()
 yb09.Show
End Sub

Private Sub mnujlfx_Click()
   Call julei
End Sub

Private Sub mnuLongData_Click()

    Call findnumberd_day
 For CountFlood = 1 To NumberNo
     Call findtimed_day(CountFlood)
     Call daytimed(LongTimeD)
     Call save_Data(LongTimeD)
  Next CountFlood
End Sub

Private Sub mnuqlch_Click()
  Call qlybch
End Sub

Private Sub mnuqlri_Click()
  Call qlybri
End Sub

Private Sub mnurihz_Click()

  Call rmxjss
End Sub

Private Sub mnurish_Click()
  Call rishow
  yb12.Show
End Sub

Private Sub mnusccl_Click()
  yb04.Show
End Sub

Private Sub mnusccy_Click()
 yb02.Show
End Sub

Private Sub mnuscry_Click()
  yb07.Show
End Sub

Private Sub mnust00_Click()
  yb03.Show
End Sub
Private Sub mnutime_Click()
  If DAindex = 1 Then
    yb01.Show
  Else
    yb05.Show
  End If
End Sub

Private Sub mnuwucd_Click()
  cscd.Show
End Sub

Private Sub mnuwuch_Click()
  csch.Show
End Sub

Private Sub mnuybda_Click()
   Call rishow
   yb26.Show
End Sub

Private Sub mnuybjy1_Click()
   Showw = "JY"
   jy = 1
   Call ybgenerajy
End Sub

Private Sub mnuybjy2_Click()
   Showw = "JY"
   jy = 2
   Call ybgenerajy
End Sub

Private Sub mnuybjy3_Click()
   Showw = "JY"
   jy = 3
   Call ybgenerajy
End Sub

Private Sub mnuybsh_Click()
   yb14.Show
End Sub

Private Sub mnuybta_Click()
  Call cishow
  yb09.Show
End Sub

Private Sub mnuycx_Click()
  Showw = "MX"
  Call ybgenera
End Sub

Private Sub mnuztxs_Click()
   yb06.Show
End Sub
