VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "洪水选择与基流分割"
   ClientHeight    =   12105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14385
   LinkTopic       =   "Form3"
   ScaleHeight     =   12105
   ScaleWidth      =   14385
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "数据输出选项"
      Height          =   1695
      Left            =   10800
      TabIndex        =   8
      Top             =   120
      Width           =   1455
      Begin VB.CheckBox Check3 
         Caption         =   "基流"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "地表径流"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "总径流"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "选择数据"
      Height          =   615
      Left            =   12600
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存场次洪水"
      Height          =   615
      Left            =   12600
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1695
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   20
      Cols            =   5
   End
   Begin VB.Frame Frame1 
      Caption         =   "操作选择"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton Option3 
         Caption         =   "基流分割"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "洪水结束选取"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "洪水开始选取"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   9975
      Left            =   120
      ScaleHeight     =   9915
      ScaleWidth      =   13995
      TabIndex        =   0
      Top             =   1920
      Width           =   14055
      Begin VB.Line Line3 
         X1              =   2280
         X2              =   8160
         Y1              =   9000
         Y2              =   7080
      End
      Begin VB.Line Line2 
         X1              =   11040
         X2              =   11040
         Y1              =   0
         Y2              =   9960
      End
      Begin VB.Line Line1 
         X1              =   1680
         X2              =   1680
         Y1              =   0
         Y2              =   9840
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

