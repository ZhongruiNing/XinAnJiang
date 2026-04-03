VERSION 5.00
Begin VB.Form yb15 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "卢氏与故县入库洪水预报选择"
   ClientHeight    =   5340
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   13800
   Icon            =   "yb15.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   13800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame7 
      Caption         =   "单位线选用规则"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   36
      Top             =   360
      Width           =   3495
      Begin VB.Label Label15 
         Caption         =   "1、单元雨强小于5毫米，选U1"
         Height          =   375
         Left            =   240
         TabIndex        =   50
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label16 
         Caption         =   "2、单元雨强小于10毫米大于5毫米,     选U2"
         Height          =   375
         Left            =   240
         TabIndex        =   49
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label17 
         Caption         =   "3、单元雨强小于15毫米大于10毫米,     选U3"
         Height          =   375
         Left            =   240
         TabIndex        =   48
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label Label18 
         Caption         =   "4、单元雨强小于20毫米大于15毫米,     选U4"
         Height          =   375
         Left            =   240
         TabIndex        =   47
         Top             =   2520
         Width           =   3015
      End
      Begin VB.Label Label19 
         Caption         =   "5、单元雨强小于25毫米大于20毫米,      选U5"
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Label Label20 
         Caption         =   "6、单元雨强大于25毫米,选U6"
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   3960
         Width           =   2895
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "产流与汇流"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3720
      TabIndex        =   17
      Top             =   1920
      Width           =   8535
      Begin VB.TextBox Text18 
         Height          =   375
         Left            =   7200
         TabIndex        =   44
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text17 
         Height          =   375
         Left            =   6288
         TabIndex        =   43
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   5380
         TabIndex        =   42
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text15 
         Height          =   375
         Left            =   4472
         TabIndex        =   41
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   3564
         TabIndex        =   40
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   2656
         TabIndex        =   39
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   1748
         TabIndex        =   38
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   840
         TabIndex        =   37
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   7200
         TabIndex        =   25
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   6288
         TabIndex        =   24
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   5380
         TabIndex        =   23
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   4472
         TabIndex        =   22
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3564
         TabIndex        =   21
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2656
         TabIndex        =   20
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1748
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   840
         TabIndex        =   18
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "单位线  选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "各块最大雨强"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "第15块"
         Height          =   255
         Left            =   7320
         TabIndex        =   33
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "第14块"
         Height          =   255
         Left            =   6408
         TabIndex        =   32
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "第13块"
         Height          =   255
         Left            =   5500
         TabIndex        =   31
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "第12块"
         Height          =   255
         Left            =   4592
         TabIndex        =   30
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "第11块"
         Height          =   255
         Left            =   3684
         TabIndex        =   29
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "第10块"
         Height          =   255
         Left            =   2776
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "第9块"
         Height          =   255
         Left            =   1868
         TabIndex        =   27
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "第8块"
         Height          =   255
         Left            =   960
         TabIndex        =   26
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "预报入库,卢氏采用"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8200
      TabIndex        =   12
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Option4 
         Caption         =   "实测"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "预报"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "卢氏实时校正选择"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10320
      TabIndex        =   9
      Top             =   120
      Width           =   1815
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "是否校正"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1935
      Left            =   12360
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
      Begin VB.CommandButton ok 
         Caption         =   "计算"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cl 
         Caption         =   "取消"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "起始流量（基流）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   720
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "故县"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "卢氏"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "预报卢氏，灵口采用"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6080
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Option1 
         Caption         =   "预报"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "实测"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   840
         Width           =   735
      End
   End
End
Attribute VB_Name = "yb15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PMax() As Single
Private Sub cl_Click()
   Unload Me
End Sub
Private Sub Form_Load()
   Dim k As Integer
   ReDim PMax(8), Unit(8)

   Call readjliu
   Call ReadPmax(PMax)
   Q00(2) = Q00(2) - Q00(1)
   If Q00(2) < 0 Then
      Q00(2) = 0
   Else
   End If
   For k = 1 To 8
     Unit(k) = "U1"
     
     If PMax(k) < 5 Then
        Unit(k) = "U1"
     Else
     End If
     
     If PMax(k) >= 5 And PMax(k) < 10 Then
        Unit(k) = "U2"
     Else
     End If
     
     If PMax(k) >= 10 And PMax(k) < 15 Then
        Unit(k) = "U3"
     Else
     End If
     
     If PMax(k) >= 15 And PMax(k) < 20 Then
        Unit(k) = "U4"
     Else
     End If
     
     If PMax(k) >= 20 And PMax(k) < 25 Then
        Unit(k) = "U5"
     Else
     End If
     
     If PMax(k) >= 25 Then
        Unit(k) = "U6"
     Else
     End If
     
   Next k
   
   Text11.Text = (CStr(Unit(1)))
   Text12.Text = (CStr(Unit(2)))
   Text13.Text = (CStr(Unit(3)))
   Text14.Text = (CStr(Unit(4)))
   Text15.Text = (CStr(Unit(5)))
   Text16.Text = (CStr(Unit(6)))
   Text17.Text = (CStr(Unit(7)))
   Text18.Text = (CStr(Unit(8)))

End Sub
Private Sub ok_Click()
Dim i As Integer
' LsLkYbQ=0  灵口预报，1 实测
' LsGxYbQ=0  卢氏预报，1 实测
' LsYbF=0    新安江模型，1 包括经验方案
' Realxz=0 不校正，1 校正
 
 If Check1 = 1 Then
   Realxz = 1
 Else
   Realxz = 0
 End If
 
 
 LkLsYbQ = 0
 If Option1 = True Then
    LkLsYbQ = 0
 Else
    LkLsYbQ = 1
 End If
 
 LsGxYbQ = 0
 If Option3 = True Then
     LsGxYbQ = 0
 Else
     LsGxYbQ = 1
 End If

 Q00(2) = Val(Text1.Text)
 Q00(3) = Val(Text2.Text)
 
 Unit(1) = Text11.Text
 Unit(2) = Text12.Text
 Unit(3) = Text13.Text
 Unit(4) = Text14.Text
 Unit(5) = Text15.Text
 Unit(6) = Text16.Text
 Unit(7) = Text17.Text
 Unit(8) = Text18.Text

 Unload Me
 
  Call jyanyubao

 End Sub

