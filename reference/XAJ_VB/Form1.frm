VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "�°���ģ�ͼ���"
   ClientHeight    =   13530
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   17925
   LinkTopic       =   "Form1"
   ScaleHeight     =   13530
   ScaleWidth      =   17925
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame16 
      Caption         =   "��ʼ����ˮ����"
      Height          =   2535
      Left            =   10800
      TabIndex        =   102
      Top             =   120
      Width           =   1455
      Begin VB.TextBox Text8 
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   105
         Text            =   "1"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   104
         Text            =   "0.5"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   103
         Text            =   "0"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "DM"
         Height          =   375
         Left            =   120
         TabIndex        =   108
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label31 
         Caption         =   "LM"
         Height          =   375
         Left            =   120
         TabIndex        =   107
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label30 
         Caption         =   "UM"
         Height          =   375
         Left            =   120
         TabIndex        =   106
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.Frame Frame15 
      Caption         =   "�չ���ģ��ʱ��"
      Height          =   1935
      Left            =   240
      TabIndex        =   79
      Top             =   720
      Width           =   3135
      Begin VB.TextBox Text6 
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   89
         Text            =   "2005-1-1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Index           =   3
         Left            =   1920
         TabIndex        =   88
         Text            =   "2006-12-31"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   83
         Text            =   "1995-1-1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Index           =   3
         Left            =   1920
         TabIndex        =   82
         Text            =   "2006-12-31"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   81
         Text            =   "1996-1-1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   3
         Left            =   1920
         TabIndex        =   80
         Text            =   "2004-12-31"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "������"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   91
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "-"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   90
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "ģ����"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   87
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "-"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   86
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "�ʶ���"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   85
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "-"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   84
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "����������"
      Height          =   10455
      Left            =   240
      TabIndex        =   63
      Top             =   2880
      Width           =   12495
      Begin VB.CommandButton Command2 
         Caption         =   "��ͼ-flood"
         Height          =   300
         Left            =   7080
         TabIndex        =   70
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   5760
         TabIndex        =   69
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��ͼ-day"
         Height          =   300
         Left            =   2280
         TabIndex        =   67
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   960
         TabIndex        =   66
         Top             =   360
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   9495
         Left            =   120
         ScaleHeight     =   9435
         ScaleWidth      =   12195
         TabIndex        =   64
         Top             =   840
         Width           =   12255
      End
      Begin VB.Label Label22 
         Caption         =   "��ˮ���̣�"
         Height          =   375
         Left            =   4800
         TabIndex        =   68
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "�չ��̣�"
         Height          =   375
         Left            =   240
         TabIndex        =   65
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "�ʶ�����ѡ��"
      Height          =   2535
      Left            =   12480
      TabIndex        =   40
      Top             =   120
      Width           =   5175
      Begin VB.Frame Frame14 
         Caption         =   "��������"
         Height          =   2175
         Left            =   4080
         TabIndex        =   76
         Top             =   240
         Width           =   975
         Begin VB.CheckBox Check1 
            Caption         =   "LR"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   78
            Top             =   960
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox Check1 
            Caption         =   "CR"
            Height          =   375
            Index           =   14
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Value           =   1  'Checked
            Width           =   495
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "�������"
         Height          =   2175
         Left            =   3000
         TabIndex        =   55
         Top             =   240
         Width           =   975
         Begin VB.CheckBox Check1 
            Caption         =   "CG"
            Height          =   375
            Index           =   13
            Left            =   120
            TabIndex        =   58
            Top             =   1320
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox Check1 
            Caption         =   "CI"
            Height          =   375
            Index           =   12
            Left            =   120
            TabIndex        =   57
            Top             =   840
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox Check1 
            Caption         =   "CS"
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Value           =   1  'Checked
            Width           =   615
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "��ˮԴ"
         Height          =   2175
         Left            =   2040
         TabIndex        =   43
         Top             =   240
         Width           =   855
         Begin VB.CheckBox Check1 
            Caption         =   "KI"
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   54
            Top             =   1680
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox Check1 
            Caption         =   "KG"
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   53
            Top             =   1320
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox Check1 
            Caption         =   "EX"
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   52
            Top             =   840
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox Check1 
            Caption         =   "SM"
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Value           =   1  'Checked
            Width           =   495
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "����"
         Height          =   2175
         Left            =   1080
         TabIndex        =   42
         Top             =   240
         Width           =   855
         Begin VB.CheckBox Check1 
            Caption         =   "IM"
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   50
            Top             =   1320
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox Check1 
            Caption         =   "B"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   49
            Top             =   840
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox Check1 
            Caption         =   "WM"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Value           =   1  'Checked
            Width           =   495
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "����"
         Height          =   2175
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   855
         Begin VB.CheckBox Check1 
            Caption         =   "C"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   47
            Top             =   1800
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox Check1 
            Caption         =   "LM"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   46
            Top             =   1320
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox Check1 
            Caption         =   "UM"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Kc"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Value           =   1  'Checked
            Width           =   615
         End
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1200
      TabIndex        =   39
      Top             =   240
      Width           =   2175
   End
   Begin VB.Frame Frame9 
      Caption         =   "ģ����"
      Height          =   10455
      Left            =   12840
      TabIndex        =   33
      Top             =   2880
      Width           =   4815
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   4095
         Left            =   120
         TabIndex        =   99
         Top             =   6240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   7223
         _Version        =   393216
         Rows            =   30
         Cols            =   4
         WordWrap        =   -1  'True
         AllowUserResizing=   3
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Index           =   1
         Left            =   3720
         TabIndex        =   95
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   94
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Index           =   1
         Left            =   2760
         TabIndex        =   93
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Index           =   0
         Left            =   2760
         TabIndex        =   92
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   62
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   59
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   35
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   34
         Top             =   600
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3615
         Left            =   120
         TabIndex        =   98
         Top             =   2040
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   50
         Cols            =   5
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label29 
         Caption         =   "Ŀ�꺯��ֵ��"
         Height          =   255
         Left            =   240
         TabIndex        =   101
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label28 
         Caption         =   "��ˮ�������̣�"
         Height          =   255
         Left            =   240
         TabIndex        =   100
         Top             =   5880
         Width           =   1335
      End
      Begin VB.Label Label27 
         Caption         =   "�ϸ���(%)"
         Height          =   255
         Left            =   3720
         TabIndex        =   97
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "���Re(%)"
         Height          =   255
         Left            =   2760
         TabIndex        =   96
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "������"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "�ʶ���"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "Ч��ϵ��"
         Height          =   255
         Left            =   840
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "����Re(%)"
         Height          =   255
         Left            =   1800
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ģ�Ͳ���ȡֵ"
      Height          =   2535
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.Frame Frame13 
         Caption         =   "��������"
         Height          =   2175
         Left            =   5520
         TabIndex        =   71
         Top             =   240
         Width           =   1335
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   15
            Left            =   480
            TabIndex        =   74
            Text            =   "1"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   14
            Left            =   480
            TabIndex        =   72
            Text            =   "0.5"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label24 
            Caption         =   "LR"
            Height          =   375
            Left            =   120
            TabIndex        =   75
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label23 
            Caption         =   "CR"
            Height          =   375
            Left            =   120
            TabIndex        =   73
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "����"
         Height          =   2175
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1215
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   28
            Text            =   "0.85"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   1
            Left            =   360
            TabIndex        =   27
            Text            =   "20"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   2
            Left            =   360
            TabIndex        =   26
            Text            =   "40"
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   3
            Left            =   360
            TabIndex        =   25
            Text            =   "0.01"
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Kc"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "UM"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "LM"
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "C"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1800
            Width           =   255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "����"
         Height          =   2175
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   1215
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   4
            Left            =   360
            TabIndex        =   20
            Text            =   "150"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   5
            Left            =   360
            TabIndex        =   19
            Text            =   "0.2"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   6
            Left            =   360
            TabIndex        =   18
            Text            =   "0.001"
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "WM"
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "B"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "IM"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   375
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "��ˮԴ"
         Height          =   2175
         Left            =   2760
         TabIndex        =   8
         Top             =   240
         Width           =   1215
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   7
            Left            =   360
            TabIndex        =   12
            Text            =   "20"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   8
            Left            =   360
            TabIndex        =   11
            Text            =   "1.2"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   9
            Left            =   360
            TabIndex        =   10
            Text            =   "0.4"
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   10
            Left            =   360
            TabIndex        =   9
            Text            =   "0.3"
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "SM"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "EX"
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label10 
            Caption         =   "KG"
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label11 
            Caption         =   "KI"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1800
            Width           =   375
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "�������"
         Height          =   2175
         Left            =   4080
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   11
            Left            =   480
            TabIndex        =   4
            Text            =   "0.1"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   12
            Left            =   480
            TabIndex        =   3
            Text            =   "0.6"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   13
            Left            =   480
            TabIndex        =   2
            Text            =   "0.99"
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "CS"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label13 
            Caption         =   "CI"
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label14 
            Caption         =   "CG"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1320
            Width           =   375
         End
      End
   End
   Begin VB.Label Label15 
      Caption         =   "����ѡ��"
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   240
      Width           =   975
   End
   Begin VB.Menu m10 
      Caption         =   "�ļ�"
      Begin VB.Menu m14 
         Caption         =   "���κ�ˮ��ѡ"
      End
      Begin VB.Menu m15 
         Caption         =   "-"
      End
      Begin VB.Menu m11 
         Caption         =   "����ͼƬ"
      End
      Begin VB.Menu m17 
         Caption         =   "�չ���ͼƬ"
      End
      Begin VB.Menu m16 
         Caption         =   "��ˮͼƬ"
      End
      Begin VB.Menu m12 
         Caption         =   "-"
      End
      Begin VB.Menu m13 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu m20 
      Caption         =   "����"
      Begin VB.Menu m21 
         Caption         =   "��ȡ-�ղ���"
      End
      Begin VB.Menu m22 
         Caption         =   "����-�ղ���"
      End
      Begin VB.Menu m25 
         Caption         =   "-"
      End
      Begin VB.Menu m23 
         Caption         =   "��ȡ-��ˮ����"
      End
      Begin VB.Menu m24 
         Caption         =   "����-��ˮ����"
      End
   End
   Begin VB.Menu m30 
      Caption         =   "����"
      Begin VB.Menu m31 
         Caption         =   "��ģ��-�����ʶ�"
      End
      Begin VB.Menu m32 
         Caption         =   "��ģ��-����ģ��"
      End
      Begin VB.Menu m35 
         Caption         =   "-"
      End
      Begin VB.Menu m33 
         Caption         =   "��ˮģ��-�����ʶ�-1h"
      End
      Begin VB.Menu m37 
         Caption         =   "-"
      End
      Begin VB.Menu m36 
         Caption         =   "10mm��λ��"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private data_day() As Single 'ԭʼ�չ������ϣ�6�У��ꡢ�¡��ա���ˮ������������
Private data_flood() As Single 'ԭʼ��ˮ�������ϣ�7�У��ꡢ�¡��ա�ʱ����ˮ������������
Private basin As String  'ѡ���Ŀ����������
Private bn() As String  '��������
Private ba() As Single  '�������


'********************����NashЧ��ϵ��********************
Private Function nce(sv() As Single, ov() As Single) As Single
    'sv:ģ��ֵ
    'ov:�۲�ֵ
    
    Dim num As Integer
    Dim ex As Single, sum As Double
    Dim temp1 As Single, temp2 As Single
    Dim i As Integer
    
    num = UBound(sv)
    For i = 1 To num
        sum = sum + ov(i)
    Next i
    ex = sum / num
    For i = 1 To num
        temp1 = temp1 + (sv(i) - ov(i)) ^ 2
        temp2 = temp2 + (ov(i) - ex) ^ 2
    Next i
    nce = 1 - temp1 / temp2
End Function

Private Function f1(a As Integer) As Integer  '�����������
    'a:���
    If a Mod 100 = 0 Then
        If a Mod 400 = 0 Then
            f1 = 366
        Else
            f1 = 365
        End If
    Else
        If a Mod 4 = 0 Then
            f1 = 366
        Else
            f1 = 365
        End If
    End If
End Function

Private Function f2(a As Integer, b As Integer) As Integer '����������
    'a:���
    'b:�·�
    Select Case b
    Case 1, 3, 5, 7, 8, 10, 12
        f2 = 31
    Case 4, 6, 9, 11
        f2 = 30
    Case 2
        If f1(a) = 365 Then
            f2 = 28
        Else
            f2 = 29
        End If
    End Select
End Function

'��1-n����Ȼ���������
Private Function ran_sample(a() As Long)
    Dim n As Long
    Dim b() As Long, x As Long
    Dim i As Long, j As Long
    
    n = UBound(a)
    ReDim Preserve b(n)
    For i = 1 To n
        b(i) = i
    Next i
    For i = 1 To n
        x = Int(Rnd * (n - i + 1)) + 1
        a(i) = b(x)
        For j = x To n - i
            b(j) = b(j + 1)
        Next j
    Next i
End Function

Private Sub Combo1_Click()
    Dim myfso As New FileSystemObject
    Dim myfolder As Folder
    Dim myfile As File

    Dim n_flood As Integer
    Dim flood_name() As String
    Dim d4 As String, c7 As Integer
    Form1.Caption = Combo1.Text & "����--�°���ģ�ͼ���"
    basin = Combo1.Text
    
    Combo3.Clear
    Set myfolder = myfso.GetFolder(App.Path & "\data\" & basin & "\�۲�����\���κ�ˮ\")
    For Each myfile In myfolder.Files
        n_flood = n_flood + 1
        ReDim Preserve flood_name(n_flood)
        d4 = myfile.Name
        c7 = InStr(1, d4, ".")
        flood_name(n_flood) = Left(d4, c7 - 1)
        Combo3.AddItem flood_name(n_flood)
    Next
    
    Combo3.Text = Combo3.List(0)
    
End Sub

'�չ��̻�ͼ
Private Sub Command1_Click()
    Dim data() As Single
    Dim data2() As Single
    Dim riqi() As Single
    Dim n As Integer
    
    Dim pqmax As Single
    Dim ppmax As Single
    Dim pn As Integer '��ͼ���г���
    Dim pp() As Single '��ͼ��ˮ����
    Dim pqo() As Single '��ͼ�۲���������
    Dim pqs() As Single '��ͼģ����������
    Dim st As Integer
    Dim dc As Single, re As Single
    Dim intep As Single
    Dim inteq As Single
    
    Dim b As String
    
    Dim i As Integer, j As Integer
    Dim a1 As String
    Dim a2 As Single
    Dim a3 As Integer, b3 As Integer
    Dim a4 As Single, b4 As Single
    
    
    Open App.Path & "\data\" & basin & "\ģ����\�չ���\�չ���ģ����.txt" For Input As #1
    Line Input #1, a1
    Do While Not EOF(1)
        n = n + 1
        ReDim Preserve data(6, n)
        For j = 1 To 6
            Input #1, data(j, n)
        Next j
        For j = 1 To 6
            Input #1, a2
        Next j
    Loop
    Close #1
    
    If Combo2.Text <> "ȫʱ��" Then
        b = "�����"
        a3 = Val(Combo2.Text)
        pn = f1(a3)
        ReDim Preserve pp(pn), pqo(pn), pqs(pn), riqi(3, pn)
        
        If a3 = data(1, 1) Then
            st = 1
        Else
            For i = data(1, 1) To a3 - 1
                st = st + f1(i)
            Next i
            st = st + 1
        End If
        For i = st To st + pn - 1
            pp(i - st + 1) = data(4, i)
            pqo(i - st + 1) = data(5, i)
            pqs(i - st + 1) = data(6, i)
            For j = 1 To 3
                riqi(j, i - st + 1) = data(j, i)
            Next j
        Next i
    Else
        b = "ȫʱ��"
        pn = n
        ReDim Preserve pp(pn), pqo(pn), pqs(pn), riqi(3, pn)
        For i = 1 To pn
            pp(i) = data(4, i)
            pqo(i) = data(5, i)
            pqs(i) = data(6, i)
            For j = 1 To 3
                riqi(j, i) = data(j, i)
            Next j

        Next i
    End If
    
    ReDim Preserve data2(3, pn)
    For i = 1 To pn
        data2(1, i) = pp(i)
        data2(2, i) = pqo(i)
        data2(3, i) = pqs(i)
    Next i
    
    Call huatu(data2, b)
    
    Picture1.CurrentX = 35
    Picture1.CurrentY = 106
    If Combo2.Text <> "ȫʱ��" Then
        Picture1.Print basin & "����" & Combo2.Text & "���ս�ˮ��������"
    Else
        Picture1.Print basin & "�����ս�ˮ��������"
    End If


    MSFlexGrid2.Rows = pn + 1
    MSFlexGrid2.ColWidth(0) = 1300
    MSFlexGrid2.RowHeight(0) = 500

    MSFlexGrid2.TextMatrix(0, 0) = "ʱ��" & Chr(13) & "(Y-M-D)"
    MSFlexGrid2.TextMatrix(0, 1) = "��ˮ��" & Chr(13) & "(mm)"
    MSFlexGrid2.TextMatrix(0, 2) = "�۲�����" & Chr(13) & "��m3/s��"
    MSFlexGrid2.TextMatrix(0, 3) = "ģ������" & Chr(13) & "��m3/s��"
    
    For i = 1 To pn
        MSFlexGrid2.TextMatrix(i, 0) = riqi(1, i) & "-" & riqi(2, i) & "-" & riqi(3, i)
        For j = 1 To 3
            MSFlexGrid2.TextMatrix(i, j) = Format(data2(j, i), "0.000")
        Next j
    Next i

End Sub

'��ˮ���̻�ͼ
Private Sub Command2_Click()
    Dim data() As Single
    Dim data2() As Single
    Dim n As Integer
    Dim fn As String
    
    Dim pqmax As Single
    Dim ppmax As Single
    Dim pn As Integer '��ͼ���г���
    Dim pp() As Single '��ͼ��ˮ����
    Dim pqo() As Single '��ͼ�۲���������
    Dim pqs() As Single '��ͼģ����������
    Dim st As Integer
    Dim dc As Single, re As Single
    
    Dim intev(7) As Integer
    
    Dim intep As Single
    Dim inteq As Single
    
    Dim i As Integer, j As Integer
    Dim a1 As String
    Dim a2 As Single
    Dim a3 As Integer, b3 As Integer
    Dim a4 As Single, b4 As Single
    Dim a5 As Integer
    
    fn = Combo3.Text
    
    Open App.Path & "\data\" & basin & "\ģ����\��ˮ����\���κ�ˮ\" & fn & ".txt" For Input As #1
    Line Input #1, a1
    Do While Not EOF(1)
        n = n + 1
        ReDim Preserve data(7, n)
        For j = 1 To 7
            Input #1, data(j, n)
        Next j
    Loop
    Close #1
    
    pn = n
    ReDim Preserve pp(pn), pqo(pn), pqs(pn), data2(3, pn)
    For i = 1 To pn
        pp(i) = data(5, i)
        pqo(i) = data(6, i)
        pqs(i) = data(7, i)
    Next i
    
    For i = 1 To pn
        data2(1, i) = pp(i)
        data2(2, i) = pqo(i)
        data2(3, i) = pqs(i)
    Next i
    
    Call huatu(data2, fn)
    
    Picture1.CurrentX = 35
    Picture1.CurrentY = 106
    Picture1.Print basin & "����" & Combo3.Text & "�ź�ˮ��ˮ��������"
    
    
    MSFlexGrid2.Rows = pn + 1
    MSFlexGrid2.ColWidth(0) = 1300
    MSFlexGrid2.RowHeight(0) = 500

    MSFlexGrid2.TextMatrix(0, 0) = "ʱ��" & Chr(13) & "(Y-M-D-H)"
    MSFlexGrid2.TextMatrix(0, 1) = "��ˮ��" & Chr(13) & "(mm)"
    MSFlexGrid2.TextMatrix(0, 2) = "�۲�����" & Chr(13) & "��m3/s��"
    MSFlexGrid2.TextMatrix(0, 3) = "ģ������" & Chr(13) & "��m3/s��"
    
    For i = 1 To pn
        MSFlexGrid2.TextMatrix(i, 0) = data(1, i) & "-" & data(2, i) & "-" & data(3, i) & "-" & data(4, i)
        For j = 1 To 3
            MSFlexGrid2.TextMatrix(i, j) = Format(data(j + 4, i), "0.000")
        Next j
    Next i
    
End Sub

Private Sub Form_Load()
    Dim myfso As New FileSystemObject
    Dim myfolder As Folder
    Dim myfile As File

    Dim n_basin As Integer
    Dim n_flood As Integer
    Dim flood_name() As String
    
    Dim stime(6) As Integer 'ģ������ֹ������

    Dim i As Integer, j As Integer
    Dim a1 As Integer, b1 As Integer
    Dim a2 As String
    Dim d4 As String, c7 As Integer
    Dim c4 As String, b7 As Integer
    Open App.Path & "\data\�����������.txt" For Input As #1
    n_basin = 0
    Do While Not EOF(1)
        n_basin = n_basin + 1
        ReDim Preserve bn(n_basin)
        ReDim Preserve ba(n_basin)
        Line Input #1, a2
        b1 = InStr(a2, Chr(9))
        bn(n_basin) = Left(a2, b1 - 1)
        ba(n_basin) = Right(a2, Len(a2) - b1)
    Loop
    Close #1
    
    For i = 1 To n_basin
        Combo1.AddItem bn(i)
    Next i
    Combo1.Text = Combo1.List(0)
    basin = Combo1.List(0)
    Form1.Caption = Combo1.Text & "����--�°���ģ�ͼ���"
    
    c4 = Text4(0).Text
    stime(1) = Val(Left(c4, 4))
    b7 = InStr(6, c4, "-")
    stime(2) = Val(Mid(c4, 6, b7 - 6))
    stime(3) = Val(Right(c4, Len(c4) - b7))
    c4 = Text4(3).Text
    stime(4) = Val(Left(c4, 4))
    b7 = InStr(6, c4, "-")
    stime(5) = Val(Mid(c4, 6, b7 - 6))
    stime(6) = Val(Right(c4, Len(c4) - b7))
    
    For i = stime(1) To stime(4)
        Combo2.AddItem i
    Next i
    Combo2.AddItem "ȫʱ��"
    Combo2.Text = Combo2.List(stime(4) - stime(1) + 1)
    
    Combo3.Clear
   
    Set myfolder = myfso.GetFolder(App.Path & "\data\" & basin & "\�۲�����\���κ�ˮ\")
    For Each myfile In myfolder.Files
        n_flood = n_flood + 1
        ReDim Preserve flood_name(n_flood)
        d4 = myfile.Name
        c7 = InStr(1, d4, ".")
        flood_name(n_flood) = Left(d4, c7 - 1)
        Combo3.AddItem flood_name(n_flood)
    Next

    Combo3.Text = Combo3.List(0)
End Sub

'����ͼƬ
Private Sub m11_Click()
     SavePicture Picture1.Image, App.Path & "\����������.bmp"
End Sub

'�˳�����
Private Sub m13_Click()
    End
End Sub

'ȷ��ģ�͵��ʶ��ںͼ�����
Private Sub m14_Click()
    Form3.Show
End Sub

'�����ˮģ��ͼƬ
Private Sub m16_Click()

    Dim data() As Single
    Dim data2() As Single
    Dim n As Integer
    Dim fn As String
    
    Dim n_flood As Integer '��ˮ���θ���
    
    Dim pqmax As Single
    Dim ppmax As Single
    Dim pn As Integer '��ͼ���г���
    Dim pp() As Single '��ͼ��ˮ����
    Dim pqo() As Single '��ͼ�۲���������
    Dim pqs() As Single '��ͼģ����������
    Dim st As Integer
    Dim dc As Single, re As Single
    
    Dim intev(7) As Integer
    
    Dim intep As Single
    Dim inteq As Single
    
    Dim i As Integer, j As Integer, h As Integer
    Dim a1 As String
    Dim a2 As Single
    Dim a3 As Integer, b3 As Integer
    Dim a4 As Single, b4 As Single
    Dim a5 As Integer
    
    n_flood = Combo3.ListCount
    For h = 1 To n_flood
        ReDim Preserve data(7, 0)
        n = 0
    
        fn = Combo3.List(h - 1)
        
        Open App.Path & "\data\" & basin & "\ģ����\��ˮ����\���κ�ˮ\" & fn & ".txt" For Input As #1
        Line Input #1, a1
        Do While Not EOF(1)
            n = n + 1
            ReDim Preserve data(7, n)
            For j = 1 To 7
                Input #1, data(j, n)
            Next j
        Loop
        Close #1
        
        ReDim Preserve data2(3, n)
        For i = 1 To n
            For j = 1 To 3
                data2(j, i) = data(j + 4, i)
            Next j
        Next i
        
        Call huatu(data2, fn)
        
        Picture1.CurrentX = 35
        Picture1.CurrentY = 106
        Picture1.Print basin & "����" & fn & "�ź�ˮ��ˮ��������"
        
        SavePicture Picture1.Image, App.Path & "\data\" & basin & "\ģ����\��ˮ����\����������\" & fn & ".bmp"
        
    Next h

MsgBox "ok"

End Sub

'�����չ���ͼƬ
Private Sub m17_Click()
    Dim data() As Single
    Dim data2() As Single
    Dim n As Integer
    Dim n_year As Integer
    
    Dim pqmax As Single
    Dim ppmax As Single
    Dim pn As Integer '��ͼ���г���
    Dim pp() As Single '��ͼ��ˮ����
    Dim pqo() As Single '��ͼ�۲���������
    Dim pqs() As Single '��ͼģ����������
    Dim st As Integer
    Dim dc As Single, re As Single
    Dim intep As Single
    Dim inteq As Single
    
    Dim b As String
    
    Dim i As Integer, j As Integer
    Dim a1 As String
    Dim a2 As Single
    Dim a3 As Integer, b3 As Integer
    Dim a4 As Single, b4 As Single
    
    
    Open App.Path & "\data\" & basin & "\ģ����\�չ���\�չ���ģ����.txt" For Input As #1

    Line Input #1, a1
    Do While Not EOF(1)
        n = n + 1
        ReDim Preserve data(6, n)
        For j = 1 To 6
            Input #1, data(j, n)
        Next j
        For j = 1 To 6
            Input #1, a2
        Next j
    Loop
    Close #1
    
    n_year = data(1, n) - data(1, 1) + 1
    
    For j = 1 To n_year
        pn = 0
        st = 0
        ReDim Preserve pp(pn), pqo(pn), pqs(pn)
        ReDim Preserve data2(3, pn)
        
        Combo2.Text = Combo2.List(j - 1)

        
        b = "�����"
        a3 = data(1, 1) + j - 1
        pn = f1(a3)
        ReDim Preserve pp(pn), pqo(pn), pqs(pn)
        
        If a3 = data(1, 1) Then
            st = 1
        Else
            For i = data(1, 1) To a3 - 1
                st = st + f1(i)
            Next i
            st = st + 1
        End If
        For i = st To st + pn - 1
            pp(i - st + 1) = data(4, i)
            pqo(i - st + 1) = data(5, i)
            pqs(i - st + 1) = data(6, i)
        Next i
    
        ReDim Preserve data2(3, pn)
        For i = 1 To pn
            data2(1, i) = pp(i)
            data2(2, i) = pqo(i)
            data2(3, i) = pqs(i)
        Next i
        
        Call huatu(data2, b)
        
        Picture1.CurrentX = 35
        Picture1.CurrentY = 106
        Picture1.Print basin & "����" & a3 & "���ս�ˮ��������"
        
        SavePicture Picture1.Image, App.Path & "\data\" & basin & "\ģ����\�չ���\����������\" & a3 & ".bmp"
    Next j
    MsgBox "OK"

End Sub

'������ģ�Ͳ���
Private Sub m21_Click()
    Dim para(16)  As Single
    Dim a4 As String
    Dim i As Integer
    Open App.Path & "\data\" & Combo1.Text & "\ģ�Ͳ���\��ģ�Ͳ�����ֵ.txt" For Input As #1
    Line Input #1, a4
    For i = 1 To 16
        Input #1, para(i)
        If para(i) < 1 Then
            Text1(i - 1).Text = Format(para(i), "0.000")
        Else
            Text1(i - 1).Text = para(i)
        End If
    Next i
    Close #1
End Sub

'�����չ��̲����ʶ����
Private Sub m22_Click()
    Dim para(16) As Single
    Dim nashc1 As Single, re1 As Single
    Dim nashc2 As Single, re2 As Single
    
    Dim i As Integer, j As Integer
    
    For i = 1 To 16
        para(i) = Text1(i - 1).Text
    Next i
    
    nashc1 = Text2(0).Text
    re1 = Text3(0).Text
    nashc2 = Text2(1).Text
    re2 = Text3(1).Text
    
    Open App.Path & "\data\" & basin & "\ģ�Ͳ���\��ģ�Ͳ�����ֵ.txt" For Output As #1
    Print #1, "Kc" & Chr(9) & "UM" & Chr(9) & "LM" & Chr(9) & "C" & Chr(9) & "WM" & Chr(9) & "B" & Chr(9) & "IM" & Chr(9) & "SM" & Chr(9) & "EX" & Chr(9) & "Kg" & Chr(9) & "Ki" & Chr(9) & "Cs" & Chr(9) & "Ci" & Chr(9) & "Cg" & Chr(9) & "Cr" & Chr(9) & "Lr"
    For i = 1 To 16
        Print #1, Format(para(i), "0.000") & Chr(9);
    Next i
    Close #1
    
    Open App.Path & "\data\" & basin & "\ģ����\�չ���\�չ��̲�����ѡ���.txt" For Output As #1
    Print #1, "Kc" & Chr(9) & "UM" & Chr(9) & "LM" & Chr(9) & "C" & Chr(9) & "WM" & Chr(9) & "B" & Chr(9) & "IM" & Chr(9) & "SM" & Chr(9) & "EX" & Chr(9) & "Kg" & Chr(9) & "Ki" & Chr(9) & "Cs" & Chr(9) & "Ci" & Chr(9) & "Cg" & Chr(9) & "Cr" & Chr(9) & "Lr"
    For i = 1 To 16
        Print #1, Format(para(i), "0.000") & Chr(9);
    Next i
    Print #1,
    Print #1,
    Print #1, "���" & Chr(9) & "DC" & Chr(9) & "Re(%)"
    Print #1, "�ʶ���" & Chr(9) & Format(nashc1, "0.000") & Chr(9) & Format(re1, "0.000")
    Print #1, "������" & Chr(9) & Format(nashc2, "0.000") & Chr(9) & Format(re2, "0.000")
    
    For i = 1 To MSFlexGrid1.Rows - 1
        For j = 0 To 2
            Print #1, MSFlexGrid1.TextMatrix(i, j) & Chr(9);
        Next j
        Print #1,
    Next i
    Close #1

End Sub

'�����ˮģ�Ͳ���
Private Sub m23_Click()
    Dim para(16)  As Single
    Dim a4 As String
    Dim i As Integer
    Open App.Path & "\data\" & Combo1.Text & "\ģ�Ͳ���\��ˮģ�Ͳ�����ֵ.txt" For Input As #1
    Line Input #1, a4
    For i = 1 To 16
        Input #1, para(i)
        If para(i) < 1 Then
            Text1(i - 1).Text = Format(para(i), "0.000")
        Else
            Text1(i - 1).Text = para(i)
        End If
    Next i
    Close #1

End Sub

'�����ˮģ�Ͳ����ʶ����
Private Sub m24_Click()
    Dim para(16) As Single
    Dim nashc1 As Single, re1 As Single
    Dim nashc2 As Single, re2 As Single
    Dim n_flood As Integer
    
    Dim i As Integer, j As Integer
    For i = 1 To 16
        para(i) = Text1(i - 1).Text
    Next i
    
    n_flood = MSFlexGrid1.Rows - 1
    nashc1 = Text2(0).Text
    re1 = Text9(0).Text
    nashc2 = Text2(1).Text
    re2 = Text9(1).Text
    
    Open App.Path & "\data\" & basin & "\ģ�Ͳ���\��ˮģ�Ͳ�����ֵ.txt" For Output As #1
    Print #1, "Kc" & Chr(9) & "UM" & Chr(9) & "LM" & Chr(9) & "C" & Chr(9) & "WM" & Chr(9) & "B" & Chr(9) & "IM" & Chr(9) & "SM" & Chr(9) & "EX" & Chr(9) & "Kg" & Chr(9) & "Ki" & Chr(9) & "Cs" & Chr(9) & "Ci" & Chr(9) & "Cg" & Chr(9) & "Cr" & Chr(9) & "Lr"
    For i = 1 To 16
        Print #1, Format(para(i), "0.000") & Chr(9);
    Next i
    Close #1
    
    
    Open App.Path & "\data\" & basin & "\ģ����\��ˮ����\��ˮ���̲����ʶ����.txt" For Output As #1
    Print #1, "Kc" & Chr(9) & "UM" & Chr(9) & "LM" & Chr(9) & "C" & Chr(9) & "WM" & Chr(9) & "B" & Chr(9) & "IM" & Chr(9) & "SM" & Chr(9) & "EX" & Chr(9) & "Kg" & Chr(9) & "Ki" & Chr(9) & "Cs" & Chr(9) & "Ci" & Chr(9) & "Cg" & Chr(9) & "Cr" & Chr(9) & "Lr"
    For i = 1 To 16
        Print #1, Format(para(i), "0.000") & Chr(9);
    Next i
    Print #1,
    Print #1,
    
    Print #1, Chr(9) & "Nsh" & Chr(9) & "�ϸ���(%)"
    Print #1, "�ʶ���" & Chr(9) & Format(nashc1, "0.000") & Chr(9) & Format(re1, "0.000")
    Print #1, "������" & Chr(9) & Format(nashc2, "0.000") & Chr(9) & Format(re2, "0.000")
    Print #1,
    
    Print #1, "flood" & Chr(9) & Chr(9) & "Nsh" & Chr(9) & "Rew(%)" & Chr(9) & "Rep(%)" & Chr(9) & "Ret"
    For i = 1 To n_flood
        For j = 0 To 4
            Print #1, MSFlexGrid1.TextMatrix(i, j) & Chr(9);
        Next j
        Print #1,
    Next i
    
    Close #1

End Sub

'��ģ�Ͳ����ʶ�
Private Sub m31_Click()
    Randomize
    
    Dim area As Single '���������km2��
    Dim u As Single '������λת������
    
    Dim p() As Single '�۲⽵ˮ
    Dim q_obs() As Single '�۲�����
    Dim epan() As Single '�۲�����������
    Dim riqi() As Integer '�۲��ꡢ�¡���
    Dim n_obs As Integer '�۲����г���
    Dim pe_obs() As Single '�۲⽵ˮ������
    
    Dim w0(3) As Single '���ʼ������ˮ�����ϡ��С��£�����
    Dim intial(3) As Single '���ʼ������ˮ�����ϡ��С��£�

    
    Dim q_sim() As Single 'ģ������
    
    Dim n_sample As Long '��������
    Dim para_in(16) As Single '������ֵ
    Dim para() As Single '�����Է����Ĳ�������
    Dim para2(16) As Single 'һ���������
    Dim para_max(16) As Single '���Ų���
    Dim num_ran() As Long '1-n_sample��Ȼ�����������
    Dim para_bound(16, 2) As Single '16��ģ�Ͳ�����������
    Dim result() As Single 'ģ����

'********************dim model parameter********************
    Dim kc As Single, um As Single, lm As Single, c As Single '����
    Dim wm As Single, b As Single, im As Single '����
    Dim sm As Single, ex As Single, kg As Single, ki As Single '��ˮԴ
    Dim cs As Single, ci As Single, cg As Single, cr As Single, lr As Integer '����
    
    Dim stime(6) As Integer 'ģ������ֹ������
    Dim ctime(6) As Integer '�ʶ�����ֹ������
    Dim vtime(6) As Integer '��������ֹ������
    Dim nashc1() As Single, re1() As Single '�ʶ���nashЧ��ϵ����������
    Dim nashc2() As Single, re2() As Single '������nashЧ��ϵ����������
    Dim ts1 As Integer, ts2 As Integer 'ģ������ֹʱ��
    Dim tc1 As Integer, tc2 As Integer '�ʶ�����ֹʱ��
    Dim tv1 As Integer, tv2 As Integer '��������ֹʱ��
    Dim nsi As Integer 'ģ���ڳ���
    Dim nca As Integer ' �ʶ��ڳ���
    Dim nve As Integer '�����ڳ���
    Dim nashc_max1 As Single '�ʶ���Ч��ϵ������ֵ
    Dim re_max1 As Single '�ʶ�������������ֵ
    Dim nashc_max2 As Single '������Ч��ϵ������ֵ
    Dim re_max2 As Single '����������������ֵ
    Dim q_sim_max() As Single '����ģ������
    Dim result_max() As Single '����ģ����
    
    Dim nsy As Integer 'ģ�����
    Dim dcy() As Single 'ģ����ÿ���ȷ����ϵ��
    Dim rey() As Single 'ģ����ÿ���������
    Dim dcy_max() As Single 'ģ����ÿ���ȷ����ϵ������ֵ
    Dim rey_max() As Single 'ģ����ÿ�������������ֵ
    Dim datay() As Single 'ģ����ÿ��Ľ�ˮ���۲⾶����ģ�⾶��������ϵ��
    
    Dim pp() As Single '��ͼ��ˮ����
    Dim pqo() As Single '��ͼ�۲���������
    Dim pqs() As Single '��ͼģ����������
    Dim data_p() As Single
    Dim pn As Integer '��ͼ���г���
    Dim b_q As String
    
    
    Dim i As Long, j As Long, temp1 As Long, k As Long, h As Long
    Dim sv() As Single, ov() As Single, pv() As Single
    Dim a2 As Single, b2 As Single, c2 As Single, d2 As Single
    Dim a3 As Integer, b3 As Integer, c3 As Integer
    Dim a4 As String, b4 As String, c4 As String
    Dim a5 As Integer, b5 As Integer, c5 As Integer
    Dim a6 As Single
    Dim a7 As Integer, b7 As Integer, c7 As Integer
    Dim a8 As Integer, b8 As Integer, c8 As Integer
    Dim a9() As Single

    
'ģ����-�ʶ���-������ȷ��
    c4 = Text4(0).Text
    stime(1) = Val(Left(c4, 4))
    b7 = InStr(6, c4, "-")
    stime(2) = Val(Mid(c4, 6, b7 - 6))
    stime(3) = Val(Right(c4, Len(c4) - b7))
    c4 = Text4(3).Text
    stime(4) = Val(Left(c4, 4))
    b7 = InStr(6, c4, "-")
    stime(5) = Val(Mid(c4, 6, b7 - 6))
    stime(6) = Val(Right(c4, Len(c4) - b7))
    
    nsy = stime(4) - stime(1) + 1
    ReDim Preserve dcy(nsy)
    ReDim Preserve rey(nsy)
    ReDim Preserve dcy_max(nsy)
    ReDim Preserve rey_max(nsy)
    ReDim Preserve datay(4, nsy), a9(nsy)
    
    c4 = Text5(0).Text
    ctime(1) = Val(Left(c4, 4))
    b7 = InStr(6, c4, "-")
    ctime(2) = Val(Mid(c4, 6, b7 - 6))
    ctime(3) = Val(Right(c4, Len(c4) - b7))
    c4 = Text5(3).Text
    ctime(4) = Val(Left(c4, 4))
    b7 = InStr(6, c4, "-")
    ctime(5) = Val(Mid(c4, 6, b7 - 6))
    ctime(6) = Val(Right(c4, Len(c4) - b7))
    
    c4 = Text6(0).Text
    vtime(1) = Val(Left(c4, 4))
    b7 = InStr(6, c4, "-")
    vtime(2) = Val(Mid(c4, 6, b7 - 6))
    vtime(3) = Val(Right(c4, Len(c4) - b7))
    c4 = Text6(3).Text
    vtime(4) = Val(Left(c4, 4))
    b7 = InStr(6, c4, "-")
    vtime(5) = Val(Mid(c4, 6, b7 - 6))
    vtime(6) = Val(Right(c4, Len(c4) - b7))

'********************ȷ���������ƺ����********************
    For i = 1 To UBound(bn)
        If basin = bn(i) Then
            area = ba(i)
            Exit For
        End If
    Next i
    u = area / 3.6 / 24

'********************���뽵ˮ����������������********************
    Open App.Path & "\data\" & basin & "\�۲�����\" & basin & "�ս�ˮ-����-����.txt" For Input As #1
    Do While Not EOF(1)
        n_obs = n_obs + 1
        ReDim Preserve p(n_obs)
        ReDim Preserve riqi(3, n_obs)
        ReDim Preserve q_obs(n_obs)
        ReDim Preserve epan(n_obs)
        For i = 1 To 3
            Input #1, riqi(i, n_obs)
        Next i
        Input #1, p(n_obs)
        Input #1, epan(n_obs)
        Input #1, q_obs(n_obs)
    Loop
    Close #1
    
    If riqi(1, 1) <> stime(1) Then
        For a7 = riqi(1, 1) To stime(1) - 1
            ts1 = ts1 + f1(a7)
        Next a7
    End If
    ts1 = ts1 + 1
    For a7 = riqi(1, 1) To stime(4)
        ts2 = ts2 + f1(a7)
    Next a7
    nsi = ts2 - ts1 + 1
    
    For a7 = riqi(1, 1) To ctime(1) - 1
        tc1 = tc1 + f1(a7)
    Next a7
    tc1 = tc1 + 1
    For a7 = riqi(1, 1) To ctime(4)
        tc2 = tc2 + f1(a7)
    Next a7
    nca = tc2 - tc1 + 1

    For a7 = riqi(1, 1) To vtime(1) - 1
        tv1 = tv1 + f1(a7)
    Next a7
    tv1 = tv1 + 1
    For a7 = riqi(1, 1) To vtime(4)
        tv2 = tv2 + f1(a7)
    Next a7
    nve = tv2 - tv1 + 1
    
    ReDim Preserve q_sim(nsi)
    ReDim Preserve pe_obs(2, nsi)
    ReDim Preserve result(7, nsi)
    ReDim Preserve q_sim_max(nsi)
    ReDim Preserve result_max(7, nsi)
    
    
    For i = ts1 To ts2
        pe_obs(1, i - ts1 + 1) = p(i)
        pe_obs(2, i - ts1 + 1) = epan(i)
    Next i
    
'********************���������Է�����������********************
    n_sample = InputBox("��������ʶ���������", , 10000)
    Open App.Path & "\data\" & basin & "\ģ�Ͳ���\��ģ�Ͳ����ʶ���Χ.txt" For Input As #1
    For i = 1 To 16
        For j = 1 To 2
            Input #1, para_bound(i, j)
        Next j
        Input #1, b4
    Next i
    Close #1
    
    ReDim Preserve para(16, n_sample), num_ran(n_sample)
    ReDim Preserve nashc1(n_sample), re1(n_sample)
    ReDim Preserve nashc2(n_sample), re2(n_sample)
    
    For i = 1 To 16
        para_in(i) = Val(Text1(i - 1).Text)
    Next i
    
    For i = 1 To 3
        w0(i) = Val(Text8(i - 1).Text)
    Next i
    
    For i = 1 To 16
        If Check1(i - 1).Value = 1 Then
            Call ran_sample(num_ran)
            For j = 1 To n_sample
                para(i, j) = para_bound(i, 1) + (para_bound(i, 2) - para_bound(i, 1)) * (num_ran(j) - 0.5) / n_sample
            Next j
        Else
            For j = 1 To n_sample
                para(i, j) = para_in(i)
            Next j
        End If
    Next i
    For i = 1 To n_sample
        If para(5, i) <= para(2, i) + para(3, i) + 20 Then
            para(5, i) = para(2, i) + para(3, i) + 20
        End If
        If para(13, i) > para(14, i) Then
            a6 = para(13, i)
            para(13, i) = para(14, i)
            para(14, i) = a6
        End If
        para(2, i) = Int(para(2, i))
        para(3, i) = Int(para(3, i))
        para(5, i) = Int(para(5, i))
        para(16, i) = Int(para(16, i))
    Next i
            

    Open App.Path & "\data\" & basin & "\ģ����\�չ���\�չ��̲����ʶ�����.txt" For Output As #2
    Print #2, "Kc" & Chr(9) & "UM" & Chr(9) & "LM" & Chr(9) & "C" & Chr(9) & "WM" & Chr(9) & "B" & Chr(9) & "IM" & Chr(9) & "SM" & Chr(9) & "EX" & Chr(9) & "Kg" & Chr(9) & "Ki" & Chr(9) & "Cs" & Chr(9) & "Ci" & Chr(9) & "Cg" & Chr(9) & "Cr" & Chr(9) & "Lr" & Chr(9) & "Nce1" & Chr(9) & "Re1" & Chr(9) & "Nce2" & Chr(9) & "Re2"

'�����Է�������
    Form2.Show
    nashc_max1 = -999
    Form2.ProgressBar1.Visible = True
    Form2.ProgressBar1.Max = n_sample
    Form2.ProgressBar1.Value = Form2.ProgressBar1.Min

    For h = 1 To n_sample
        For i = 1 To 16
            para2(i) = para(i, h)
        Next i
        
        intial(1) = w0(1) * para2(2)
        intial(2) = w0(2) * para2(3)
        intial(3) = w0(3) * (para2(5) - para2(2) - para2(3))
        
        Call xaj_day(para2, pe_obs, result, intial)
        For i = 1 To nsi
            q_sim(i) = result(1, i) * u
        Next i
        
        For i = 1 To nsy
            a8 = stime(1) + i - 1
            b8 = f1(a8)
            c8 = 0
            ReDim Preserve sv(b8), ov(b8), pv(b8)
            If a8 <> riqi(1, 1) Then
                For a7 = riqi(1, 1) To a8 - 1
                    c8 = c8 + f1(a7)
                Next a7
            End If
            c8 = c8 + 1
            a2 = 0: b2 = 0: d2 = 0
            For j = c8 To c8 + b8 - 1
                ov(j - c8 + 1) = q_obs(j)
                sv(j - c8 + 1) = q_sim(j - ts1 + 1)
                pv(j - c8 + 1) = p(j)
                a2 = a2 + ov(j - c8 + 1)
                b2 = b2 + sv(j - c8 + 1)
                d2 = d2 + pv(j - c8 + 1)
            Next j
            dcy(i) = nce(sv, ov)
            rey(i) = (b2 - a2) / a2 * 100
            
            datay(1, i) = d2
            datay(2, i) = a2 / u
            a9(i) = b2 / u
            datay(4, i) = datay(2, i) / datay(1, i)
        Next i
        
        ReDim Preserve sv(nca), ov(nca)
        a2 = 0: b2 = 0
        For i = tc1 To tc2
            ov(i - tc1 + 1) = q_obs(i)
            sv(i - tc1 + 1) = q_sim(i - ts1 + 1)
            a2 = a2 + ov(i - tc1 + 1)
            b2 = b2 + sv(i - tc1 + 1)
        Next i
        nashc1(h) = nce(sv, ov)
        re1(h) = (b2 - a2) / a2 * 100
            
        ReDim Preserve sv(nve), ov(nve)
        a2 = 0: b2 = 0
        For i = tv1 To tv2
            ov(i - tv1 + 1) = q_obs(i)
            sv(i - tv1 + 1) = q_sim(i - ts1 + 1)
            a2 = a2 + ov(i - tv1 + 1)
            b2 = b2 + sv(i - tv1 + 1)
        Next i
        nashc2(h) = nce(sv, ov)
        re2(h) = (b2 - a2) / a2 * 100
    
        If nashc1(h) > nashc_max1 Then
            nashc_max1 = nashc1(h)
            re_max1 = re1(h)
            nashc_max2 = nashc2(h)
            re_max2 = re2(h)
            For i = 1 To nsy
                dcy_max(i) = dcy(i)
                rey_max(i) = rey(i)
                datay(3, i) = a9(i)
            Next i
            For i = 1 To 16
                para_max(i) = para2(i)
            Next i
            For i = 1 To nsi
                q_sim_max(i) = q_sim(i)
                For j = 2 To 7
                    result_max(j, i) = result(j, i)
                Next j
                result_max(1, i) = result(1, i) * u
            Next i
        End If
    
        For i = 1 To 16
            Print #2, Format(para2(i), "0.000") & Chr(9);
        Next i
        Print #2, Format(nashc1(h), "0.000") & Chr(9);
        Print #2, Format(re1(h), "0.000") & Chr(9);
        Print #2, Format(nashc2(h), "0.000") & Chr(9);
        Print #2, Format(re2(h), "0.000")
        
        Form2.ProgressBar1.Value = h
    Next h
    Close #2
    Form2.Hide
    
'������Ų���ֵ
    For i = 1 To 16
        Text1(i - 1).Text = Format(para_max(i), "0.000")
    Next i
    Text2(0).Text = Format(nashc_max1, "0.000")
    Text3(0).Text = Format(re_max1, "0.000")
    Text2(1).Text = Format(nashc_max2, "0.000")
    Text3(1).Text = Format(re_max2, "0.000")
    
    MSFlexGrid1.Cols = 7
    MSFlexGrid1.Rows = nsy + 1
    MSFlexGrid1.RowHeight(0) = 500

    MSFlexGrid1.ColWidth(1) = 1000
    MSFlexGrid1.ColWidth(2) = 1000

    MSFlexGrid1.TextMatrix(0, 0) = "���"
    MSFlexGrid1.TextMatrix(0, 1) = "Ч��ϵ��"
    MSFlexGrid1.TextMatrix(0, 2) = "������" & Chr(13) & "��%��"
    MSFlexGrid1.TextMatrix(0, 3) = "��ˮ��" & Chr(13) & "(mm)"
    MSFlexGrid1.TextMatrix(0, 4) = "�۲⾶��" & Chr(13) & "��mm��"
    MSFlexGrid1.TextMatrix(0, 5) = "ģ�⾶��" & Chr(13) & "��mm��"
    MSFlexGrid1.TextMatrix(0, 6) = "����ϵ��"
    
    For i = 1 To nsy
        MSFlexGrid1.TextMatrix(i, 0) = stime(1) + i - 1
        MSFlexGrid1.TextMatrix(i, 1) = Format(dcy_max(i), "0.000")
        MSFlexGrid1.TextMatrix(i, 2) = Format(rey_max(i), "0.00")
        For j = 1 To 4
            MSFlexGrid1.TextMatrix(i, 2 + j) = Format(datay(j, i), "0.00")
        Next j
        
    Next i
    
    MSFlexGrid2.Rows = nsi + 1
    MSFlexGrid2.ColWidth(0) = 1300
    MSFlexGrid2.RowHeight(0) = 500

    MSFlexGrid2.TextMatrix(0, 0) = "ʱ��" & Chr(13) & "(Y-M-D)"
    MSFlexGrid2.TextMatrix(0, 1) = "��ˮ��" & Chr(13) & "(mm)"
    MSFlexGrid2.TextMatrix(0, 2) = "�۲�����" & Chr(13) & "��m3/s��"
    MSFlexGrid2.TextMatrix(0, 3) = "ģ������" & Chr(13) & "��m3/s��"
    
    For i = ts1 To ts2
        MSFlexGrid2.TextMatrix(i - ts1 + 1, 0) = riqi(1, i) & "-" & riqi(2, i) & "-" & riqi(3, i)
        MSFlexGrid2.TextMatrix(i - ts1 + 1, 1) = Format(p(i), "0.000")
        MSFlexGrid2.TextMatrix(i - ts1 + 1, 2) = Format(q_obs(i), "0.000")
        MSFlexGrid2.TextMatrix(i - ts1 + 1, 3) = Format(q_sim_max(i - ts1 + 1), "0.000")
    Next i

'���������ѡ���
    Open App.Path & "\data\" & basin & "\ģ����\�չ���\�չ��̲�����ѡ���.txt" For Output As #1
    Print #1, "Kc" & Chr(9) & "UM" & Chr(9) & "LM" & Chr(9) & "C" & Chr(9) & "WM" & Chr(9) & "B" & Chr(9) & "IM" & Chr(9) & "SM" & Chr(9) & "EX" & Chr(9) & "Kg" & Chr(9) & "Ki" & Chr(9) & "Cs" & Chr(9) & "Ci" & Chr(9) & "Cg" & Chr(9) & "Cr" & Chr(9) & "Lr"
    For i = 1 To 16
        Print #1, Format(para_max(i), "0.000") & Chr(9);
    Next i
    Print #1,
    Print #1,
    Print #1, "���" & Chr(9) & "DC" & Chr(9) & "Re(%)" & Chr(9) & "p(mm)" & Chr(9) & "R-o(mm)" & Chr(9) & "R-s(mm)" & Chr(9) & "RC"
    Print #1, "�ʶ���" & Chr(9) & Format(nashc_max1, "0.000") & Chr(9) & Format(re_max1, "0.000")
    Print #1, "������" & Chr(9) & Format(nashc_max2, "0.000") & Chr(9) & Format(re_max2, "0.000")
    For i = 1 To nsy
        Print #1, stime(1) + i - 1 & Chr(9) & Format(dcy_max(i), "0.000") & Chr(9) & Format(rey_max(i), "0.00") & Chr(9);
        For j = 1 To 3
            Print #1, Format(datay(j, i), "0.0") & Chr(9);
        Next j
        Print #1, Format(datay(4, i), "0.000")
    Next i
    Close #1

'�����������ģ����
    Open App.Path & "\data\" & basin & "\ģ����\�չ���\�չ���ģ����.txt" For Output As #1
    Print #1, "Y" & Chr(9) & "M" & Chr(9) & "D" & Chr(9) & "P" & Chr(9) & "Q-obs" & Chr(9) & "Q-sim" & Chr(9) & "E" & Chr(9) & "WU" & Chr(9) & "WL" & Chr(9) & "WD" & Chr(9) & "W" & Chr(9) & "S"
    For i = ts1 To ts2
        For j = 1 To 3
            Print #1, riqi(j, i) & Chr(9);
        Next j
        If p(i) < 1 And p(i) > 0 Then
            Print #1, Format(p(i), "0.000") & Chr(9);
        Else
            Print #1, p(i) & Chr(9);
        End If
        Print #1, q_obs(i) & Chr(9);
        
        Print #1, Format(q_sim_max(i - ts1 + 1), "0.000") & Chr(9);
        For j = 2 To 7
            Print #1, Format(result_max(j, i - ts1 + 1), "0.000") & Chr(9);
        Next j
        Print #1,
    Next i
    Close #1
    
'������������
    pn = nsi
    b_q = "ȫʱ��"
    ReDim Preserve data_p(3, pn)
    For i = ts1 To ts2
        data_p(1, i - ts1 + 1) = p(i)
        data_p(2, i - ts1 + 1) = q_obs(i)
        data_p(3, i - ts1 + 1) = q_sim_max(i - ts1 + 1)
    Next i
    Call huatu(data_p, b_q)
    
    Picture1.CurrentX = 35
    Picture1.CurrentY = 106
    Picture1.Print basin & "�����ս�ˮ��������"
    
    MsgBox "�������"
    
End Sub

'��ģ�͹���ģ��
Private Sub m32_Click()
    Randomize
    
    Dim area As Single '���������km2��
    Dim u As Single '������λת������
    
    Dim p() As Single '�۲⽵ˮ
    Dim q_obs() As Single '�۲�����
    Dim epan() As Single '�۲�����������
    Dim riqi() As Integer '�۲��ꡢ�¡���
    Dim n_obs As Integer '�۲����г���
    Dim pe_obs() As Single '�۲⽵ˮ������
    
    Dim w0(3) As Single '���ʼ������ˮ�����ϡ��С��£�����
    Dim intial(3) As Single '���ʼ������ˮ�����ϡ��С��£�

    
    Dim q_sim() As Single 'ģ������
    
    Dim n_sample As Long '��������
    Dim para_in(16) As Single '������ֵ
    Dim para() As Single '�����Է����Ĳ�������
    Dim para2(16) As Single 'һ���������
    Dim para_max(16) As Single '���Ų���
    Dim num_ran() As Long '1-n_sample��Ȼ�����������
    Dim para_bound(16, 2) As Single '16��ģ�Ͳ�����������
    Dim result() As Single '����ģ����
    Dim result_final() As Single 'ȫ��ģ����

'********************dim model parameter********************
    Dim kc As Single, um As Single, lm As Single, c As Single '����
    Dim wm As Single, b As Single, im As Single '����
    Dim sm As Single, ex As Single, kg As Single, ki As Single '��ˮԴ
    Dim cs As Single, ci As Single, cg As Single, cr As Single, lr As Integer '����
    
    Dim stime(6) As Integer 'ģ������ֹ������
    Dim ctime(6) As Integer '�ʶ�����ֹ������
    Dim vtime(6) As Integer '��������ֹ������
    Dim nashc1() As Single, re1() As Single '�ʶ���nashЧ��ϵ����������
    Dim nashc2() As Single, re2() As Single '������nashЧ��ϵ����������
    Dim ts1 As Integer, ts2 As Integer 'ģ������ֹʱ��
    Dim tc1 As Integer, tc2 As Integer '�ʶ�����ֹʱ��
    Dim tv1 As Integer, tv2 As Integer '��������ֹʱ��
    Dim nsi As Integer 'ģ���ڳ���
    Dim nca As Integer ' �ʶ��ڳ���
    Dim nve As Integer '�����ڳ���
    Dim nashc_max1 As Single '�ʶ���Ч��ϵ������ֵ
    Dim re_max1 As Single '�ʶ�������������ֵ
    Dim nashc_max2 As Single '������Ч��ϵ������ֵ
    Dim re_max2 As Single '����������������ֵ
    Dim q_sim_max() As Single '����ģ������
    Dim result_max() As Single '����ģ����
    
    Dim nsy As Integer 'ģ�����
    Dim dcy() As Single 'ģ����ÿ���ȷ����ϵ��
    Dim rey() As Single 'ģ����ÿ���������
    Dim dcy_max() As Single 'ģ����ÿ���ȷ����ϵ������ֵ
    Dim rey_max() As Single 'ģ����ÿ�������������ֵ
    Dim datay() As Single 'ģ����ÿ��Ľ�ˮ���۲⾶����ģ�⾶��������ϵ��
    
    Dim pp() As Single '��ͼ��ˮ����
    Dim pqo() As Single '��ͼ�۲���������
    Dim pqs() As Single '��ͼģ����������
    Dim data_p() As Single
    Dim pn As Integer '��ͼ���г���
    Dim b_q As String
    
    Dim n_sa As Integer 'ģ���ʱ�θ���
    Dim num_sa() As Integer 'ÿ��ʱ�εĳ���
    Dim weizhi() As Integer 'ÿ��ʱ�ε���ֹλ��
    
    
    Dim i As Long, j As Long, temp1 As Long, k As Long, h As Long
    Dim sv() As Single, ov() As Single, pv() As Single
    Dim a2 As Single, b2 As Single, c2 As Single, d2 As Single
    Dim a3 As Integer, b3 As Integer, c3 As Integer
    Dim a4 As String, b4 As String, c4 As String
    Dim a5 As Integer, b5 As Integer, c5 As Integer
    Dim a6 As Single
    Dim a7 As Integer, b7 As Integer, c7 As Integer
    Dim a8 As Integer, b8 As Integer, c8 As Integer
    Dim a9() As Single
    Dim a10 As Boolean
   

'********************ȷ���������ƺ����********************
    For i = 1 To UBound(bn)
        If basin = bn(i) Then
            area = ba(i)
            Exit For
        End If
    Next i
    u = area / 3.6 / 24
    
'����ģ�Ͳ���
    For i = 1 To 16
        para2(i) = Val(Text1(i - 1).Text)
    Next i
    For i = 1 To 3
        w0(i) = Val(Text8(i - 1).Text)
    Next i


'********************���뽵ˮ����������������********************
    Open App.Path & "\data\" & basin & "\�۲�����\" & basin & "�ս�ˮ-����-����.txt" For Input As #1
    Do While Not EOF(1)
        n_obs = n_obs + 1
        ReDim Preserve riqi(3, n_obs)
        ReDim Preserve p(n_obs)
        ReDim Preserve epan(n_obs)
        ReDim Preserve q_obs(n_obs)
        For i = 1 To 3
            Input #1, riqi(i, n_obs)
        Next i
        Input #1, p(n_obs)
        Input #1, epan(n_obs)
        Input #1, q_obs(n_obs)
    Loop
    Close #1
    ReDim Preserve result_final(7, n_obs)
    For i = 1 To n_obs
        For j = 1 To 7
            result_final(j, i) = -99
        Next j
    Next i

'ȷ������ʱ��
    a10 = False
    For i = 1 To n_obs
        If a10 = False Then
            If p(i) >= 0 Then
                n_sa = n_sa + 1
                ReDim Preserve weizhi(2, n_sa)
                weizhi(1, n_sa) = i
                a10 = True
            End If
        Else
            If p(i) < 0 Then
                weizhi(2, n_sa) = i - 1
                a10 = False
            End If
        End If
    Next i
    If weizhi(2, n_sa) = 0 Then
        weizhi(2, n_sa) = n_obs
    End If
    ReDim Preserve num_sa(n_sa)
    
'ģ�ͼ���
    For h = 1 To n_sa
        num_sa(h) = weizhi(2, h) - weizhi(1, h) + 1
        ReDim Preserve pe_obs(2, num_sa(h))
        ReDim Preserve result(7, num_sa(h))
        For i = weizhi(1, h) To weizhi(2, h)
            pe_obs(1, i - weizhi(1, h) + 1) = p(i)
            pe_obs(2, i - weizhi(1, h) + 1) = epan(i)
        Next i
        
        intial(1) = w0(1) * para2(2)
        intial(2) = w0(2) * para2(3)
        intial(3) = w0(3) * (para2(5) - para2(2) - para2(3))
        
        
        Call xaj_day(para2, pe_obs, result, intial)
        For i = weizhi(1, h) To weizhi(2, h)
            For j = 1 To 7
                result_final(j, i) = result(j, i - weizhi(1, h) + 1)
            Next j
        Next i
    Next h

'���ģ����
    Open App.Path & "\data\" & basin & "\ģ����\�չ���\�չ���ģ����.txt" For Output As #1
    Print #1, "Y" & Chr(9) & "M" & Chr(9) & "D" & Chr(9) & "P" & Chr(9) & "Q-obs" & Chr(9) & "Q-sim" & Chr(9) & "E" & Chr(9) & "WU" & Chr(9) & "WL" & Chr(9) & "WD" & Chr(9) & "W" & Chr(9) & "S"
    For i = 1 To n_obs
        For j = 1 To 3
            Print #1, riqi(j, i) & Chr(9);
        Next j
        If p(i) < 1 And p(i) > 0 Then
            Print #1, Format(p(i), "0.000") & Chr(9);
        Else
            Print #1, p(i) & Chr(9);
        End If
        Print #1, q_obs(i) & Chr(9);
        If result_final(1, i) > 0 Then
            Print #1, Format(result_final(1, i) * u, "0.000") & Chr(9);
        Else
            Print #1, Format(result_final(1, i), "0") & Chr(9);
        End If
        For j = 2 To 7
            If result_final(j, i) > 0 Then
                Print #1, Format(result_final(j, i), "0.000") & Chr(9);
            Else
                Print #1, Format(result_final(j, i), "0") & Chr(9);
            End If
        Next j
        Print #1,
    Next i
    Close #1
    
    MsgBox "�������"
   
End Sub

'��ˮģ�Ͳ����ʶ�
Private Sub m33_Click()
 Randomize
    
    Dim myfso As New FileSystemObject
    Dim myfolder As Folder
    Dim myfile As File
    
    Dim area As Single '���������km2��
    Dim u As Single '������λת������
    
    Dim flood_name() As String '��ˮ���
    Dim n_flood As Integer '��ˮ����
    Dim nfc As Integer '�ʶ��ں�ˮ����
    Dim nfv As Integer '�����ں�ˮ����
    Dim ns() As Integer 'ÿ����ˮ��ʱ�γ�
    Dim pqf() As Single '���κ�ˮ�۲⽵ˮ��������
    Dim w0() As Single 'ÿ����ˮ�ĳ�ʼ������ˮ�����ϡ��С��¡��ܣ�������ˮ��ˮ��
    Dim intial(6) As Single '��ʼ������������ˮ�����ϡ��С��¡��ܣ�������ˮ��ˮ��
    
    Dim n_epan As Integer '������ϵ�г���
    Dim n_daysim As Integer '�չ���ģ��ϵ�г���
    Dim epan_day() As Single '�۲�������������
    
    Dim p() As Single '�۲⽵ˮ
    Dim q_obs() As Single '�۲�����
    Dim riqi() As Integer '�۲��ꡢ�¡���
    Dim n_obs As Integer '�۲����г���
    Dim pe_obs() As Single '�۲⽵ˮ������
    
    Dim q_sim() As Single 'ģ������
    Dim q_sim_n() As Single
    
    Dim n_sample As Long '��������
    Dim para_in(16) As Single '������ֵ
    Dim para() As Single '�����Է����Ĳ�������
    Dim para2(16) As Single 'һ���������
    Dim para_max(16) As Single '���Ų���
    Dim num_ran() As Long '1-n_sample��Ȼ�����������
    Dim para_bound(16, 2) As Single '16��ģ�Ͳ�����������
    Dim result() As Single 'ģ����
    Dim result_day() As Single
    Dim para_day(16) As Single '��ģ�͵�16������ֵ

'********************dim model parameter********************
    Dim kc As Single, um As Single, lm As Single, c As Single '����
    Dim wm As Single, b As Single, im As Single '����
    Dim sm As Single, ex As Single, kg As Single, ki As Single '��ˮԴ
    Dim ci As Single, cg As Single, cr As Single, lr As Integer '����
    
    Dim nashc() As Single, rew() As Single, rep() As Single, ret() As Single '�ʶ���nashЧ��ϵ����������
    Dim obva1() As Single '�ʶ���ƽ��Ŀ�꺯����Ч��ϵ������������������������ʱ�����
    Dim obva2() As Single '������ƽ��Ŀ�꺯����Ч��ϵ������������������������ʱ�����
    Dim obva_n() As Single 'ÿ����ˮ��Ŀ�꺯��
    Dim obva_n_max() As Single 'ÿ����ˮ������Ŀ�꺯��
    Dim tc1 As Integer, tc2 As Integer '�ʶ�����ֹʱ��
    Dim tv1 As Integer, tv2 As Integer '��������ֹʱ��
    Dim nca As Integer ' �ʶ��ڳ���
    Dim nve As Integer '�����ڳ���
    Dim nashc_max1 As Single '�ʶ���Ч��ϵ������ֵ
    Dim re_max1 As Single '�ʶ�������������ֵ
    Dim nashc_max2 As Single '������Ч��ϵ������ֵ
    Dim re_max2 As Single '����������������ֵ
    Dim q_sim_max() As Single '����ģ������
    Dim result_max() As Single '����ģ����
    Dim obva_max1(4) As Single
    Dim obva_max2(4) As Single
    Dim dataf() As Single 'ÿ����ˮ�Ľ�ˮ���۲⾶����ģ�⾶��������ϵ�����۲��塢ģ���塢�۲����ʱ�䡢ģ�����ʱ��

    
    Dim qr1() As Single '�ʶ��ںϸ���
    Dim qr2() As Single '�����ںϸ���
    Dim qr1_max As Single
    Dim qr2_max As Single
    
    Dim pqmax As Single
    Dim ppmax As Single
    Dim pp() As Single '��ͼ��ˮ����
    Dim pqo() As Single '��ͼ�۲���������
    Dim pqs() As Single '��ͼģ����������
    Dim pn As Integer
    Dim fn As String
    Dim data_p() As Single
    
    Dim i As Long, j As Long, temp1 As Long, k As Long, h As Long
    Dim sv() As Single, ov() As Single
    Dim a2 As Single, b2 As Single, c2 As Single
    Dim a3 As Integer, b3 As Integer, c3 As Integer
    Dim a4 As String, b4 As String, c4 As String, d4 As String
    Dim a5 As Integer, b5 As Integer, c5 As Integer
    Dim a6 As Single, b6 As Single
    Dim a7 As Integer, b7 As Integer, c7 As Integer
    Dim a8 As Single, b8 As Single
    Dim a9 As Integer, b9 As Integer
    Dim a10 As Single
    Dim a11() As Single
    Dim a12 As String
    

'********************ȷ���������ƺ����********************
    For i = 1 To UBound(bn)
        If basin = bn(i) Then
            area = ba(i)
            Exit For
        End If
    Next i
    u = area / 3.6

'********************���뽵ˮ����������������********************
    Set myfolder = myfso.GetFolder(App.Path & "\data\" & basin & "\�۲�����\���κ�ˮ\")
    For Each myfile In myfolder.Files
        n_flood = n_flood + 1
        ReDim Preserve flood_name(n_flood)
        d4 = myfile.Name
        c7 = InStr(1, d4, ".")
        flood_name(n_flood) = Left(d4, c7 - 1)
    Next
    ReDim Preserve ns(n_flood)
    ReDim Preserve dataf(8, n_flood)

    
'    nfc = Int(n_flood * 0.8)
'    nfv = n_flood - nfc

    nfc = n_flood
    nfv = Int(n_flood * 0.2)
    If nfv = 0 Then nfv = 1

    
    b7 = 0
    For i = 1 To n_flood
        Open App.Path & "\data\" & basin & "\�۲�����\���κ�ˮ\" & flood_name(i) & ".txt" For Input As #1
        Do While Not EOF(1)
            ns(i) = ns(i) + 1
            If ns(i) > b7 Then
                b7 = ns(i)
                ReDim Preserve pqf(n_flood, 7, ns(i))
            End If
            For j = 1 To 6
                Input #1, pqf(i, j, ns(i))
            Next j
        Loop
        Close #1
    Next i
    
    ReDim Preserve q_sim_max(n_flood, b7)
    ReDim Preserve q_sim_n(n_flood, b7)

    
    Open App.Path & "\data\" & basin & "\�۲�����\" & basin & "�ս�ˮ-����-����.txt" For Input As #1
    Do While Not EOF(1)
        n_epan = n_epan + 1
        ReDim Preserve epan_day(4, n_epan)
        For i = 1 To 3
            Input #1, epan_day(i, n_epan)
        Next i
        Input #1, a10
        Input #1, epan_day(4, n_epan)
        Input #1, a10
    Loop
    Close #1
    
    
    Open App.Path & "\data\" & basin & "\ģ����\�չ���\�չ���ģ����.txt" For Input As #1
    Line Input #1, c4
    Do While Not EOF(1)
        n_daysim = n_daysim + 1
        ReDim Preserve result_day(12, n_daysim)
        For i = 1 To 12
            Input #1, result_day(i, n_daysim)
        Next i
    Loop
    Close #1
    
    ReDim Preserve w0(5, n_flood)
    For i = 1 To n_flood
        For j = 1 To ns(i)
            If pqf(i, 4, j) >= 6 And pqf(i, 4, j) <= 19 Then
                For k = 1 To n_epan
                    If pqf(i, 1, j) = epan_day(1, k) And pqf(i, 2, j) = epan_day(2, k) And pqf(i, 3, j) = epan_day(3, k) Then
                        If epan_day(4, k) > 0 Then
                            pqf(i, 7, j) = epan_day(4, k) / 14
                        Else
                            pqf(i, 7, j) = 0.1
                        End If
                        GoTo 12
                    End If
                Next k
12:
            Else
                pqf(i, 7, j) = 0
            End If
        Next j
    Next i
    
    For i = 1 To n_flood
        For k = 1 To n_daysim
            If pqf(i, 1, 1) = result_day(1, k) And pqf(i, 2, 1) = result_day(2, k) And pqf(i, 3, 1) = result_day(3, k) Then
                If pqf(i, 4, 1) > 7 Then
                    For j = 1 To 5
                        w0(j, i) = result_day(7 + j, k - 1)
                    Next j
                Else
                    For j = 1 To 5
                        w0(j, i) = result_day(7 + j, k - 2)
                    Next j
                End If
                GoTo 11
            End If
        Next k
11:
    Next i

    ReDim Preserve q_obs(n_obs)
    ReDim Preserve result(3, n_obs)
    ReDim Preserve result_max(3, n_obs)
    
'��ȡ��ģ�Ͳ����ʶ����
    Open App.Path & "\data\" & basin & "\ģ����\�չ���\�չ��̲�����ѡ���.txt" For Input As #1
    Line Input #1, a12
    For i = 1 To 16
        Input #1, para_day(i)
    Next i
    Close #1
    
'********************���������Է�����������********************
    n_sample = InputBox("��������ʶ���������", , 10000)
    Open App.Path & "\data\" & basin & "\ģ�Ͳ���\��ˮģ�Ͳ����ʶ���Χ.txt" For Input As #1
    For i = 1 To 16
        For j = 1 To 2
            Input #1, para_bound(i, j)
        Next j
        Input #1, b4
    Next i
    Close #1

    ReDim Preserve para(16, n_sample), num_ran(n_sample)
    ReDim Preserve obva1(4, n_sample), obva2(4, n_sample)
    ReDim Preserve nashc(n_flood), rew(n_flood), rep(n_flood), ret(n_flood)
    ReDim Preserve obva_n(n_flood, 4), obva_n_max(n_flood, 4)
    ReDim Preserve qr1(n_sample), qr2(n_sample)
    
    
    For i = 1 To 16
        para_in(i) = Val(Text1(i - 1).Text)
    Next i
    
    For i = 1 To 16
        If Check1(i - 1).Value = 1 Then
            Call ran_sample(num_ran)
            For j = 1 To n_sample
                para(i, j) = para_bound(i, 1) + (para_bound(i, 2) - para_bound(i, 1)) * (num_ran(j) - 0.5) / n_sample
            Next j
        Else
            For j = 1 To n_sample
                para(i, j) = para_in(i)
            Next j
        End If
    Next i
    For i = 1 To n_sample
        If para(5, i) <= para(2, i) + para(3, i) + 20 Then
            para(5, i) = para(2, i) + para(3, i) + 20
        End If
        If para(12, i) > para(13, i) Then
            a6 = para(12, i)
            para(12, i) = para(13, i)
            para(13, i) = a6
        End If
        para(2, i) = Int(para(2, i))
        para(3, i) = Int(para(3, i))
        para(5, i) = Int(para(5, i))
        para(16, i) = Int(para(16, i))
    Next i
            

    Open App.Path & "\data\" & basin & "\ģ����\��ˮ����\��ˮ���̲����ʶ�����.txt" For Output As #2
    Print #2, "Kc" & Chr(9) & "UM" & Chr(9) & "LM" & Chr(9) & "C" & Chr(9) & "WM" & Chr(9) & "B" & Chr(9) & "IM" & Chr(9) & "SM" & Chr(9) & "EX" & Chr(9) & "Kg" & Chr(9) & "Ki" & Chr(9) & "Cs" & Chr(9) & "Ci" & Chr(9) & "Cg" & Chr(9) & "Cr" & Chr(9) & "Lr" & Chr(9) & "Nce1" & Chr(9) & "Rew1" & Chr(9) & "Rep1" & Chr(9) & "Ret1" & Chr(9) & "Nce2" & Chr(9) & "Rew2" & Chr(9) & "Rep2" & Chr(9) & "Ret2"

'�����Է�������
    Form2.Show
    nashc_max = 0
    Form2.ProgressBar1.Visible = True
    Form2.ProgressBar1.Max = n_sample
    Form2.ProgressBar1.Value = Form2.ProgressBar1.Min
    
    obva_max1(1) = -99

    For h = 1 To n_sample
        For i = 1 To 16
            para2(i) = para(i, h)
        Next i
        For k = 1 To n_flood
            ReDim Preserve pe_obs(2, ns(k))
            ReDim Preserve q_sim(ns(k))
            
            For i = 1 To ns(k)
                pe_obs(1, i) = pqf(k, 5, i)
                pe_obs(2, i) = pqf(k, 7, i)
                q_sim(i) = 0
            Next i
            
            intial(1) = pqf(k, 6, 1)
            For i = 1 To 4
                intial(1 + i) = w0(i, k)
            Next i
            If w0(5, k) <> 0 Then
                intial(6) = para2(8) / para_day(8) * w0(5, k)
            Else
                intial(6) = 0
            End If
        
            Call xaj_flood(para2, pe_obs, q_sim, intial, u)
            
            ReDim Preserve sv(ns(k)), ov(ns(k))
            a2 = 0: b2 = 0
            a8 = 0: b8 = 0
            For i = 1 To ns(k)
                ov(i) = pqf(k, 6, i)
                sv(i) = q_sim(i)
                a2 = a2 + ov(i)
                b2 = b2 + sv(i)
                If ov(i) > a8 Then
                    a8 = ov(i)
                    a9 = i
                End If
                If sv(i) > b8 Then
                    b8 = sv(i)
                    b9 = i
                End If
                q_sim_n(k, i) = q_sim(i)
            Next i
            nashc(k) = nce(sv, ov)
            rew(k) = (b2 - a2) / a2 * 100
            rep(k) = (b8 - a8) / a8 * 100
            ret(k) = b9 - a9
        Next k
        
        For k = 1 To n_flood
            obva_n(k, 1) = nashc(k)
            obva_n(k, 2) = rew(k)
            obva_n(k, 3) = rep(k)
            obva_n(k, 4) = ret(k)
        Next k
        
        For k = 1 To nfc
            obva1(1, h) = obva1(1, h) + nashc(k)
            obva1(2, h) = obva1(2, h) + Abs(rew(k))
            obva1(3, h) = obva1(2, h) + Abs(rep(k))
            obva1(4, h) = obva1(2, h) + Abs(ret(k))
            If Abs(rew(k)) <= 20 And Abs(rep(k)) <= 20 Then
                qr1(h) = qr1(h) + 1 / nfc
            End If
        Next k
        qr1(h) = qr1(h) * 100
        
        For k = 1 + n_flood - nfv To n_flood
            obva2(1, h) = obva2(1, h) + nashc(k)
            obva2(2, h) = obva2(2, h) + Abs(rew(k))
            obva2(3, h) = obva2(2, h) + Abs(rep(k))
            obva2(4, h) = obva2(2, h) + Abs(ret(k))
            If Abs(rew(k)) <= 20 And Abs(rep(k)) <= 20 Then
                qr2(h) = qr2(h) + 1 / nfv
            End If
        Next k
        qr2(h) = qr2(h) * 100
        
        For j = 1 To 4
            obva1(j, h) = obva1(j, h) / nfc
            obva2(j, h) = obva2(j, h) / nfv
        Next j
        
        If (qr1(h) > qr1_max) Or ((qr1(h) = qr1_max) And (obva1(1, h) > obva_max1(1))) Then
            qr1_max = qr1(h)
            qr2_max = qr2(h)
            For j = 1 To 4
                obva_max1(j) = obva1(j, h)
                obva_max2(j) = obva2(j, h)
            Next j
            For i = 1 To n_flood
                For j = 1 To 4
                    obva_n_max(i, j) = obva_n(i, j)
                Next j
            Next i
            For i = 1 To 16
                para_max(i) = para2(i)
            Next i
            
            For i = 1 To n_flood
                For j = 1 To ns(i)
                    q_sim_max(i, j) = q_sim_n(i, j)
                Next j
            Next i
        End If
        
        For i = 1 To 16
            Print #2, Format(para2(i), "0.000") & Chr(9);
        Next i
        For i = 1 To 4
            Print #2, Format(obva1(i, h), "0.000") & Chr(9);
        Next i
        For i = 1 To 4
            Print #2, Format(obva2(i, h), "0.000") & Chr(9);
        Next i
        Print #2,
        Form2.ProgressBar1.Value = h
    Next h
    Close #2
    Form2.Hide
    
'������Ų���ֵ

    For i = 1 To 16
        Text1(i - 1).Text = Format(para_max(i), "0.000")
    Next i
    
    Text2(0).Text = Format(obva_max1(1), "0.000")
    Text3(0).Text = Format(obva_max1(2), "0.000")
    Text7(0).Text = Format(obva_max1(3), "0.000")
    Text9(0).Text = Format(qr1_max, "0.00")
    
    Text2(1).Text = Format(obva_max2(1), "0.000")
    Text3(1).Text = Format(obva_max2(2), "0.000")
    Text7(1).Text = Format(obva_max2(3), "0.000")
    Text9(1).Text = Format(qr2_max, "0.00")
    
    
    MSFlexGrid1.Cols = 9
    MSFlexGrid1.Rows = n_flood + 1
    MSFlexGrid1.RowHeight(0) = 500

    MSFlexGrid1.ColWidth(0) = 1000
    MSFlexGrid1.ColWidth(1) = 700
    MSFlexGrid1.ColWidth(2) = 900
    MSFlexGrid1.ColWidth(3) = 900
    MSFlexGrid1.ColWidth(4) = 1200
    
    MSFlexGrid1.TextMatrix(0, 0) = "��ˮ���"
    MSFlexGrid1.TextMatrix(0, 1) = "Ч��" & Chr(13) & "ϵ��"
    MSFlexGrid1.TextMatrix(0, 2) = "�������" & Chr(13) & "��%��"
    MSFlexGrid1.TextMatrix(0, 3) = "������" & Chr(13) & "��%��"
    MSFlexGrid1.TextMatrix(0, 4) = "����ʱ���" & Chr(13) & "��h��"
    
    MSFlexGrid1.TextMatrix(0, 5) = "��ˮ��" & Chr(13) & "(mm)"
    MSFlexGrid1.TextMatrix(0, 6) = "�۲⾶��" & Chr(13) & "��mm��"
    MSFlexGrid1.TextMatrix(0, 7) = "ģ�⾶��" & Chr(13) & "��mm��"
    MSFlexGrid1.TextMatrix(0, 8) = "����ϵ��"


    For i = 1 To n_flood
        MSFlexGrid1.TextMatrix(i, 0) = flood_name(i)
        For j = 1 To 3
            MSFlexGrid1.TextMatrix(i, j) = Format(obva_n_max(i, j), "0.000")
        Next j
        MSFlexGrid1.TextMatrix(i, 4) = Format(obva_n_max(i, 4), "#0")
        
        For j = 1 To ns(i)
            dataf(1, i) = dataf(1, i) + pqf(i, 5, j)
            dataf(2, i) = dataf(2, i) + pqf(i, 6, j) / u
            dataf(3, i) = dataf(3, i) + q_sim_max(i, j) / u
        Next j
        dataf(4, i) = dataf(2, i) / dataf(1, i)
        
        dataf(5, i) = 0
        dataf(6, i) = 0
        
        For k = 1 To ns(i)
            If dataf(5, i) < pqf(i, 6, k) Then
                dataf(5, i) = pqf(i, 6, k)
                dataf(7, i) = k
            End If
            If dataf(6, i) < q_sim_max(i, k) Then
                dataf(6, i) = q_sim_max(i, k)
                dataf(8, i) = k
            End If
        Next k
        
        
        For j = 1 To 4
            MSFlexGrid1.TextMatrix(i, 4 + j) = Format(dataf(j, i), "0.00")
        Next j
    Next i
    
    
    MSFlexGrid2.Rows = ns(1) + 1
    MSFlexGrid2.ColWidth(0) = 1300
    MSFlexGrid2.RowHeight(0) = 500

    MSFlexGrid2.TextMatrix(0, 0) = "ʱ��" & Chr(13) & "(Y-M-D-H)"
    MSFlexGrid2.TextMatrix(0, 1) = "��ˮ��" & Chr(13) & "(mm)"
    MSFlexGrid2.TextMatrix(0, 2) = "�۲�����" & Chr(13) & "��m3/s��"
    MSFlexGrid2.TextMatrix(0, 3) = "ģ������" & Chr(13) & "��m3/s��"
    
    For i = 1 To ns(1)
        MSFlexGrid2.TextMatrix(i, 0) = pqf(1, 1, i) & "-" & pqf(1, 2, i) & "-" & pqf(1, 3, i) & "-" & pqf(1, 4, i)
        For j = 1 To 2
            MSFlexGrid2.TextMatrix(i, j) = Format(pqf(1, j + 4, i), "0.000")
        Next j
        MSFlexGrid2.TextMatrix(i, j) = Format(q_sim_max(1, i), "0.000")
    Next i


'�����������ģ����
    For i = 1 To n_flood
        Open App.Path & "\data\" & basin & "\ģ����\��ˮ����\���κ�ˮ\" & flood_name(i) & ".txt" For Output As #1
        Print #1, "Y" & Chr(9) & "M" & Chr(9) & "D" & Chr(9) & "T" & Chr(9) & "P" & Chr(9) & "Q-obs" & Chr(9) & "Q-sim"
        For k = 1 To ns(i)
            For j = 1 To 6
                If j <= 4 Then
                    Print #1, pqf(i, j, k) & Chr(9);
                Else
                    If pqf(i, j, k) = 0 Then
                        Print #1, pqf(i, j, k) & Chr(9);
                    Else
                        Print #1, Format(pqf(i, j, k), "0.000") & Chr(9);
                    End If
                End If
            Next j
            Print #1, Format(q_sim_max(i, k), "0.000")
        Next k
        Close #1
    Next i
    
    Open App.Path & "\data\" & basin & "\ģ����\��ˮ����\��ˮ���̲����ʶ����.txt" For Output As #1
    Print #1, "Kc" & Chr(9) & "UM" & Chr(9) & "LM" & Chr(9) & "C" & Chr(9) & "WM" & Chr(9) & "B" & Chr(9) & "IM" & Chr(9) & "SM" & Chr(9) & "EX" & Chr(9) & "Kg" & Chr(9) & "Ki" & Chr(9) & "Cs" & Chr(9) & "Ci" & Chr(9) & "Cg" & Chr(9) & "Cr" & Chr(9) & "Lr"
    For i = 1 To 16
        Print #1, Format(para_max(i), "0.000") & Chr(9);
    Next i
    Print #1,
    Print #1,
    
    Print #1, Chr(9) & "Nsh" & Chr(9) & "�ϸ���(%)"
    Print #1, "�ʶ���" & Chr(9) & Format(obva_max1(1), "0.000") & Chr(9) & Format(qr1_max, "0.000")
    Print #1, "������" & Chr(9) & Format(obva_max2(1), "0.000") & Chr(9) & Format(qr2_max, "0.000")
    Print #1,
    
    Print #1, "flood" & Chr(9) & "Nsh" & Chr(9) & "Rew(%)" & Chr(9) & "Rep(%)" & Chr(9) & "Ret" & Chr(9) & "p(mm)" & Chr(9) & "R-o(mm)" & Chr(9) & "R-s(mm)" & Chr(9) & "RC" & Chr(9) & "P-o(m3/s)" & Chr(9) & "P-s(m3/s)" & Chr(9) & "T-o(h)" & Chr(9) & "T-s(h)"
    For i = 1 To n_flood
        Print #1, flood_name(i) & Chr(9);
        Print #1, Format(obva_n_max(i, 1), "0.000") & Chr(9);
        
        For j = 2 To 3
            Print #1, Format(obva_n_max(i, j), "0.00") & Chr(9);
        Next j
        Print #1, Format(obva_n_max(i, 4), "#0") & Chr(9);
        For j = 1 To 3
            Print #1, Format(dataf(j, i), "0.0") & Chr(9);
        Next j
        Print #1, Format(dataf(4, i), "0.000") & Chr(9);
        For j = 5 To 6
            Print #1, Format(dataf(j, i), "0.00") & Chr(9);
        Next j
        For j = 7 To 8
            Print #1, Format(dataf(j, i), "#0") & Chr(9);
        Next j
        Print #1,
    Next i
    Close #1
    
    fn = flood_name(1)
    pn = ns(1)
    ReDim Preserve pp(pn), pqo(pn), pqs(pn), data_p(3, pn)
    For i = 1 To pn
        For j = 1 To 2
            data_p(j, i) = pqf(1, j + 4, i)
        Next j
        data_p(3, i) = q_sim_max(1, i)
    Next i
    
    Call huatu(data_p, fn)
    
    Picture1.CurrentX = 35
    Picture1.CurrentY = 106
    Picture1.Print basin & "����" & fn & "�ź�ˮ��ˮ��������"
    
    MsgBox "�������"
    
End Sub

'��ˮ�����ʶ�-30����
Private Sub m34_Click()
    Randomize
    
    Dim myfso As New FileSystemObject
    Dim myfolder As Folder
    Dim myfile As File
    
    Dim area As Single '���������km2��
    Dim u As Single '������λת������
    
    Dim flood_name() As String '��ˮ���
    Dim n_flood As Integer '��ˮ����
    Dim nfc As Integer '�ʶ��ں�ˮ����
    Dim nfv As Integer '�����ں�ˮ����
    Dim ns() As Integer 'ÿ����ˮ��ʱ�γ�
    Dim pqf() As Single '���κ�ˮ�۲⽵ˮ��������
    Dim w0() As Single 'ÿ����ˮ�ĳ�ʼ������ˮ�����ϡ��С��¡��ܣ�������ˮ��ˮ��
    Dim intial(6) As Single '��ʼ������������ˮ�����ϡ��С��¡��ܣ�������ˮ��ˮ��
    
    Dim n_epan As Integer '������ϵ�г���
    Dim n_daysim As Integer '�չ���ģ��ϵ�г���
    Dim epan_day() As Single '�۲�������������
    
    Dim p() As Single '�۲⽵ˮ
    Dim q_obs() As Single '�۲�����
    Dim riqi() As Integer '�۲��ꡢ�¡���
    Dim n_obs As Integer '�۲����г���
    Dim pe_obs() As Single '�۲⽵ˮ������
    
    Dim q_sim() As Single 'ģ������
    Dim q_sim_n() As Single
    
    Dim n_sample As Long '��������
    Dim para_in(16) As Single '������ֵ
    Dim para() As Single '�����Է����Ĳ�������
    Dim para2(16) As Single 'һ���������
    Dim para_max(16) As Single '���Ų���
    Dim num_ran() As Long '1-n_sample��Ȼ�����������
    Dim para_bound(16, 2) As Single '16��ģ�Ͳ�����������
    Dim result() As Single 'ģ����
    Dim result_day() As Single
    Dim para_day(16) As Single '��ģ�͵�16������ֵ

'********************dim model parameter********************
    Dim kc As Single, um As Single, lm As Single, c As Single '����
    Dim wm As Single, b As Single, im As Single '����
    Dim sm As Single, ex As Single, kg As Single, ki As Single '��ˮԴ
    Dim ci As Single, cg As Single, cr As Single, lr As Integer '����
    
    Dim nashc() As Single, rew() As Single, rep() As Single, ret() As Single '�ʶ���nashЧ��ϵ����������
    Dim obva1() As Single '�ʶ���ƽ��Ŀ�꺯����Ч��ϵ������������������������ʱ�����
    Dim obva2() As Single '������ƽ��Ŀ�꺯����Ч��ϵ������������������������ʱ�����
    Dim obva_n() As Single 'ÿ����ˮ��Ŀ�꺯��
    Dim obva_n_max() As Single 'ÿ����ˮ������Ŀ�꺯��
    Dim tc1 As Integer, tc2 As Integer '�ʶ�����ֹʱ��
    Dim tv1 As Integer, tv2 As Integer '��������ֹʱ��
    Dim nca As Integer ' �ʶ��ڳ���
    Dim nve As Integer '�����ڳ���
    Dim nashc_max1 As Single '�ʶ���Ч��ϵ������ֵ
    Dim re_max1 As Single '�ʶ�������������ֵ
    Dim nashc_max2 As Single '������Ч��ϵ������ֵ
    Dim re_max2 As Single '����������������ֵ
    Dim q_sim_max() As Single '����ģ������
    Dim result_max() As Single '����ģ����
    Dim obva_max1(4) As Single
    Dim obva_max2(4) As Single
    Dim dataf() As Single 'ÿ����ˮ�Ľ�ˮ���۲⾶����ģ�⾶��������ϵ��

    
    Dim qr1() As Single '�ʶ��ںϸ���
    Dim qr2() As Single '�����ںϸ���
    Dim qr1_max As Single
    Dim qr2_max As Single
    
    Dim pqmax As Single
    Dim ppmax As Single
    Dim pp() As Single '��ͼ��ˮ����
    Dim pqo() As Single '��ͼ�۲���������
    Dim pqs() As Single '��ͼģ����������
    Dim pn As Integer
    Dim fn As String
    Dim data_p() As Single
    
    Dim i As Long, j As Long, temp1 As Long, k As Long, h As Long
    Dim sv() As Single, ov() As Single
    Dim a2 As Single, b2 As Single, c2 As Single
    Dim a3 As Integer, b3 As Integer, c3 As Integer
    Dim a4 As String, b4 As String, c4 As String, d4 As String
    Dim a5 As Integer, b5 As Integer, c5 As Integer
    Dim a6 As Single, b6 As Single
    Dim a7 As Integer, b7 As Integer, c7 As Integer
    Dim a8 As Single, b8 As Single
    Dim a9 As Integer, b9 As Integer
    Dim a10 As Single
    Dim a11() As Single
    Dim a12 As String
    

'********************ȷ���������ƺ����********************
    For i = 1 To UBound(bn)
        If basin = bn(i) Then
            area = ba(i)
            Exit For
        End If
    Next i
    u = area / 1.8

'********************���뽵ˮ����������������********************
    Set myfolder = myfso.GetFolder(App.Path & "\data\" & basin & "\�۲�����\���κ�ˮ\")
    For Each myfile In myfolder.Files
        n_flood = n_flood + 1
        ReDim Preserve flood_name(n_flood)
        d4 = myfile.Name
        c7 = InStr(1, d4, ".")
        flood_name(n_flood) = Left(d4, c7 - 1)
    Next
    ReDim Preserve ns(n_flood)
    ReDim Preserve dataf(4, n_flood)

    
'    nfc = Int(n_flood * 0.8)
'    nfv = n_flood - nfc

    nfc = n_flood
    nfv = Int(n_flood * 0.2)
    If nfv = 0 Then nfv = 1

    
    b7 = 0
    For i = 1 To n_flood
        Open App.Path & "\data\" & basin & "\�۲�����\���κ�ˮ\" & flood_name(i) & ".txt" For Input As #1
        Do While Not EOF(1)
            ns(i) = ns(i) + 1
            If ns(i) > b7 Then
                b7 = ns(i)
                ReDim Preserve pqf(n_flood, 7, ns(i))
            End If
            For j = 1 To 6
                Input #1, pqf(i, j, ns(i))
            Next j
        Loop
        Close #1
    Next i
    
    ReDim Preserve q_sim_max(n_flood, b7)
    ReDim Preserve q_sim_n(n_flood, b7)

    
    Open App.Path & "\data\" & basin & "\�۲�����\" & basin & "�ս�ˮ-����-����.txt" For Input As #1
    Do While Not EOF(1)
        n_epan = n_epan + 1
        ReDim Preserve epan_day(4, n_epan)
        For i = 1 To 3
            Input #1, epan_day(i, n_epan)
        Next i
        Input #1, a10
        Input #1, epan_day(4, n_epan)
        Input #1, a10
    Loop
    Close #1
    
    
    Open App.Path & "\data\" & basin & "\ģ����\�չ���\�չ���ģ����.txt" For Input As #1
    Line Input #1, c4
    Do While Not EOF(1)
        n_daysim = n_daysim + 1
        ReDim Preserve result_day(12, n_daysim)
        For i = 1 To 12
            Input #1, result_day(i, n_daysim)
        Next i
    Loop
    Close #1
    
    ReDim Preserve w0(5, n_flood)
    For i = 1 To n_flood
        For j = 1 To ns(i)
            If pqf(i, 4, j) >= 6 And pqf(i, 4, j) <= 19 Then
                For k = 1 To n_epan
                    If pqf(i, 1, j) = epan_day(1, k) And pqf(i, 2, j) = epan_day(2, k) And pqf(i, 3, j) = epan_day(3, k) Then
                        pqf(i, 7, j) = epan_day(4, k) / 28
                        GoTo 12
                    End If
                Next k
12:
            Else
                pqf(i, 7, j) = 0
            End If
        Next j
    Next i
    
    For i = 1 To n_flood
        For k = 1 To n_daysim
            If pqf(i, 1, 1) = result_day(1, k) And pqf(i, 2, 1) = result_day(2, k) And pqf(i, 3, 1) = result_day(3, k) Then
                If pqf(i, 4, 1) > 7 Then
                    For j = 1 To 5
                        w0(j, i) = result_day(7 + j, k - 1)
                    Next j
                Else
                    For j = 1 To 5
                        w0(j, i) = result_day(7 + j, k - 2)
                    Next j
                End If
                GoTo 11
            End If
        Next k
11:
    Next i

    ReDim Preserve q_obs(n_obs)
    ReDim Preserve result(3, n_obs)
    ReDim Preserve result_max(3, n_obs)
    
'��ȡ��ģ�Ͳ����ʶ����
    Open App.Path & "\data\" & basin & "\ģ����\�չ���\�չ��̲�����ѡ���.txt" For Input As #1
    Line Input #1, a12
    For i = 1 To 16
        Input #1, para_day(i)
    Next i
    Close #1
    
'********************���������Է�����������********************
    n_sample = InputBox("��������ʶ���������", , 10000)
    Open App.Path & "\data\" & basin & "\ģ�Ͳ���\��ˮģ�Ͳ����ʶ���Χ.txt" For Input As #1
    For i = 1 To 16
        For j = 1 To 2
            Input #1, para_bound(i, j)
        Next j
        Input #1, b4
    Next i
    Close #1

    ReDim Preserve para(16, n_sample), num_ran(n_sample)
    ReDim Preserve obva1(4, n_sample), obva2(4, n_sample)
    ReDim Preserve nashc(n_flood), rew(n_flood), rep(n_flood), ret(n_flood)
    ReDim Preserve obva_n(n_flood, 4), obva_n_max(n_flood, 4)
    ReDim Preserve qr1(n_sample), qr2(n_sample)
    
    
    For i = 1 To 16
        para_in(i) = Val(Text1(i - 1).Text)
    Next i
    
    For i = 1 To 16
        If Check1(i - 1).Value = 1 Then
            Call ran_sample(num_ran)
            For j = 1 To n_sample
                para(i, j) = para_bound(i, 1) + (para_bound(i, 2) - para_bound(i, 1)) * (num_ran(j) - 0.5) / n_sample
            Next j
        Else
            For j = 1 To n_sample
                para(i, j) = para_in(i)
            Next j
        End If
    Next i
    For i = 1 To n_sample
        If para(5, i) <= para(2, i) + para(3, i) + 20 Then
            para(5, i) = para(2, i) + para(3, i) + 20
        End If
        If para(12, i) > para(13, i) Then
            a6 = para(12, i)
            para(12, i) = para(13, i)
            para(13, i) = a6
        End If
        para(2, i) = Int(para(2, i))
        para(3, i) = Int(para(3, i))
        para(5, i) = Int(para(5, i))
        para(16, i) = Int(para(16, i))
    Next i
            

    Open App.Path & "\data\" & basin & "\ģ����\��ˮ����\��ˮ���̲����ʶ�����.txt" For Output As #2
    Print #2, "Kc" & Chr(9) & "UM" & Chr(9) & "LM" & Chr(9) & "C" & Chr(9) & "WM" & Chr(9) & "B" & Chr(9) & "IM" & Chr(9) & "SM" & Chr(9) & "EX" & Chr(9) & "Kg" & Chr(9) & "Ki" & Chr(9) & "Cs" & Chr(9) & "Ci" & Chr(9) & "Cg" & Chr(9) & "Cr" & Chr(9) & "Lr" & Chr(9) & "Nce1" & Chr(9) & "Rew1" & Chr(9) & "Rep1" & Chr(9) & "Ret1" & Chr(9) & "Nce2" & Chr(9) & "Rew2" & Chr(9) & "Rep2" & Chr(9) & "Ret2"

'�����Է�������
    Form2.Show
    nashc_max = 0
    Form2.ProgressBar1.Visible = True
    Form2.ProgressBar1.Max = n_sample
    Form2.ProgressBar1.Value = Form2.ProgressBar1.Min
    
    obva_max1(1) = -99

    For h = 1 To n_sample
        For i = 1 To 16
            para2(i) = para(i, h)
        Next i
        For k = 1 To n_flood
            ReDim Preserve pe_obs(2, ns(k))
            ReDim Preserve q_sim(ns(k))
            
            For i = 1 To ns(k)
                pe_obs(1, i) = pqf(k, 5, i)
                pe_obs(2, i) = pqf(k, 7, i)
                q_sim(i) = 0
            Next i
            
            intial(1) = pqf(k, 6, 1)
            For i = 1 To 4
                intial(1 + i) = w0(i, k)
            Next i
            If w0(5, k) <> 0 Then
                intial(6) = para2(8) / para_day(8) * w0(5, k)
            Else
                intial(6) = 0
            End If
        
            Call xaj_flood_min(para2, pe_obs, q_sim, intial, u)
            
            ReDim Preserve sv(ns(k)), ov(ns(k))
            a2 = 0: b2 = 0
            a8 = 0: b8 = 0
            For i = 1 To ns(k)
                ov(i) = pqf(k, 6, i)
                sv(i) = q_sim(i)
                a2 = a2 + ov(i)
                b2 = b2 + sv(i)
                If ov(i) > a8 Then
                    a8 = ov(i)
                    a9 = i
                End If
                If sv(i) > b8 Then
                    b8 = sv(i)
                    b9 = i
                End If
                q_sim_n(k, i) = q_sim(i)
            Next i
            nashc(k) = nce(sv, ov)
            rew(k) = (b2 - a2) / a2 * 100
            rep(k) = (b8 - a8) / a8 * 100
            ret(k) = b9 - a9
        Next k
        
        For k = 1 To n_flood
            obva_n(k, 1) = nashc(k)
            obva_n(k, 2) = rew(k)
            obva_n(k, 3) = rep(k)
            obva_n(k, 4) = ret(k)
        Next k
        
        For k = 1 To nfc
            obva1(1, h) = obva1(1, h) + nashc(k)
            obva1(2, h) = obva1(2, h) + Abs(rew(k))
            obva1(3, h) = obva1(2, h) + Abs(rep(k))
            obva1(4, h) = obva1(2, h) + Abs(ret(k))
            If Abs(rew(k)) <= 20 And Abs(rep(k)) <= 20 Then
                qr1(h) = qr1(h) + 1 / nfc
            End If
        Next k
        qr1(h) = qr1(h) * 100
        
        For k = 1 + n_flood - nfv To n_flood
            obva2(1, h) = obva2(1, h) + nashc(k)
            obva2(2, h) = obva2(2, h) + Abs(rew(k))
            obva2(3, h) = obva2(2, h) + Abs(rep(k))
            obva2(4, h) = obva2(2, h) + Abs(ret(k))
            If Abs(rew(k)) <= 20 And Abs(rep(k)) <= 20 Then
                qr2(h) = qr2(h) + 1 / nfv
            End If
        Next k
        qr2(h) = qr2(h) * 100
        
        For j = 1 To 4
            obva1(j, h) = obva1(j, h) / nfc
            obva2(j, h) = obva2(j, h) / nfv
        Next j
        
        If (qr1(h) > qr1_max) Or ((qr1(h) = qr1_max) And (obva1(1, h) > obva_max1(1))) Then
            qr1_max = qr1(h)
            qr2_max = qr2(h)
            For j = 1 To 4
                obva_max1(j) = obva1(j, h)
                obva_max2(j) = obva2(j, h)
            Next j
            For i = 1 To n_flood
                For j = 1 To 4
                    obva_n_max(i, j) = obva_n(i, j)
                Next j
            Next i
            For i = 1 To 16
                para_max(i) = para2(i)
            Next i
            
            For i = 1 To n_flood
                For j = 1 To ns(i)
                    q_sim_max(i, j) = q_sim_n(i, j)
                Next j
            Next i
        End If
        
        For i = 1 To 16
            Print #2, Format(para2(i), "0.000") & Chr(9);
        Next i
        For i = 1 To 4
            Print #2, Format(obva1(i, h), "0.000") & Chr(9);
        Next i
        For i = 1 To 4
            Print #2, Format(obva2(i, h), "0.000") & Chr(9);
        Next i
        Print #2,
        Form2.ProgressBar1.Value = h
    Next h
    Close #2
    Form2.Hide
    
'������Ų���ֵ
        



    For i = 1 To 16
        Text1(i - 1).Text = Format(para_max(i), "0.000")
    Next i
    
    Text2(0).Text = Format(obva_max1(1), "0.000")
    Text3(0).Text = Format(obva_max1(2), "0.000")
    Text7(0).Text = Format(obva_max1(3), "0.000")
    Text9(0).Text = Format(qr1_max, "0.00")
    
    Text2(1).Text = Format(obva_max2(1), "0.000")
    Text3(1).Text = Format(obva_max2(2), "0.000")
    Text7(1).Text = Format(obva_max2(3), "0.000")
    Text9(1).Text = Format(qr2_max, "0.00")
    
    
    MSFlexGrid1.Cols = 9
    MSFlexGrid1.Rows = n_flood + 1
    MSFlexGrid1.RowHeight(0) = 500

    MSFlexGrid1.ColWidth(0) = 1000
    MSFlexGrid1.ColWidth(1) = 700
    MSFlexGrid1.ColWidth(2) = 900
    MSFlexGrid1.ColWidth(3) = 900
    MSFlexGrid1.ColWidth(4) = 1200
    
    MSFlexGrid1.TextMatrix(0, 0) = "��ˮ���"
    MSFlexGrid1.TextMatrix(0, 1) = "Ч��" & Chr(13) & "ϵ��"
    MSFlexGrid1.TextMatrix(0, 2) = "�������" & Chr(13) & "��%��"
    MSFlexGrid1.TextMatrix(0, 3) = "������" & Chr(13) & "��%��"
    MSFlexGrid1.TextMatrix(0, 4) = "����ʱ���" & Chr(13) & "��h��"
    
    MSFlexGrid1.TextMatrix(0, 5) = "��ˮ��" & Chr(13) & "(mm)"
    MSFlexGrid1.TextMatrix(0, 6) = "�۲⾶��" & Chr(13) & "��mm��"
    MSFlexGrid1.TextMatrix(0, 7) = "ģ�⾶��" & Chr(13) & "��mm��"
    MSFlexGrid1.TextMatrix(0, 8) = "����ϵ��"


    For i = 1 To n_flood
        MSFlexGrid1.TextMatrix(i, 0) = flood_name(i)
        For j = 1 To 3
            MSFlexGrid1.TextMatrix(i, j) = Format(obva_n_max(i, j), "0.000")
        Next j
        MSFlexGrid1.TextMatrix(i, 4) = Format(obva_n_max(i, 4), "#0")
        
        For j = 1 To ns(i)
            dataf(1, i) = dataf(1, i) + pqf(i, 5, j)
            dataf(2, i) = dataf(2, i) + pqf(i, 6, j) / u
            dataf(3, i) = dataf(3, i) + q_sim_max(i, j) / u
        Next j
        dataf(4, i) = dataf(2, i) / dataf(1, i)
        
        For j = 1 To 4
            MSFlexGrid1.TextMatrix(i, 4 + j) = Format(dataf(j, i), "0.00")
        Next j
    Next i
    
    
    MSFlexGrid2.Rows = ns(1) + 1
    MSFlexGrid2.ColWidth(0) = 1300
    MSFlexGrid2.RowHeight(0) = 500

    MSFlexGrid2.TextMatrix(0, 0) = "ʱ��" & Chr(13) & "(Y-M-D-H)"
    MSFlexGrid2.TextMatrix(0, 1) = "��ˮ��" & Chr(13) & "(mm)"
    MSFlexGrid2.TextMatrix(0, 2) = "�۲�����" & Chr(13) & "��m3/s��"
    MSFlexGrid2.TextMatrix(0, 3) = "ģ������" & Chr(13) & "��m3/s��"
    
    For i = 1 To ns(1)
        MSFlexGrid2.TextMatrix(i, 0) = pqf(1, 1, i) & "-" & pqf(1, 2, i) & "-" & pqf(1, 3, i) & "-" & pqf(1, 4, i)
        For j = 1 To 2
            MSFlexGrid2.TextMatrix(i, j) = Format(pqf(1, j + 4, i), "0.000")
        Next j
        MSFlexGrid2.TextMatrix(i, j) = Format(q_sim_max(1, i), "0.000")
    Next i


'�����������ģ����
    For i = 1 To n_flood
        Open App.Path & "\data\" & basin & "\ģ����\��ˮ����\���κ�ˮ\" & flood_name(i) & ".txt" For Output As #1
        Print #1, "Y" & Chr(9) & "M" & Chr(9) & "D" & Chr(9) & "T" & Chr(9) & "P" & Chr(9) & "Q-obs" & Chr(9) & "Q-sim"
        For k = 1 To ns(i)
            For j = 1 To 6
                If j <= 4 Then
                    Print #1, pqf(i, j, k) & Chr(9);
                Else
                    If pqf(i, j, k) = 0 Then
                        Print #1, pqf(i, j, k) & Chr(9);
                    Else
                        Print #1, Format(pqf(i, j, k), "0.000") & Chr(9);
                    End If
                End If
            Next j
            Print #1, Format(q_sim_max(i, k), "0.000")
        Next k
        Close #1
    Next i
    
    Open App.Path & "\data\" & basin & "\ģ����\��ˮ����\��ˮ���̲����ʶ����.txt" For Output As #1
    Print #1, "Kc" & Chr(9) & "UM" & Chr(9) & "LM" & Chr(9) & "C" & Chr(9) & "WM" & Chr(9) & "B" & Chr(9) & "IM" & Chr(9) & "SM" & Chr(9) & "EX" & Chr(9) & "Kg" & Chr(9) & "Ki" & Chr(9) & "Cs" & Chr(9) & "Ci" & Chr(9) & "Cg" & Chr(9) & "Cr" & Chr(9) & "Lr"
    For i = 1 To 16
        Print #1, Format(para_max(i), "0.000") & Chr(9);
    Next i
    Print #1,
    Print #1,
    
    Print #1, Chr(9) & "Nsh" & Chr(9) & "�ϸ���(%)"
    Print #1, "�ʶ���" & Chr(9) & Format(obva_max1(1), "0.000") & Chr(9) & Format(qr1_max, "0.000")
    Print #1, "������" & Chr(9) & Format(obva_max2(1), "0.000") & Chr(9) & Format(qr2_max, "0.000")
    Print #1,
    
    Print #1, "flood" & Chr(9) & Chr(9) & "Nsh" & Chr(9) & "Rew(%)" & Chr(9) & "Rep(%)" & Chr(9) & "Ret" & Chr(9) & "p(mm)" & Chr(9) & "R-o(mm)" & Chr(9) & "R-s(mm)" & Chr(9) & "RC"
    For i = 1 To n_flood
        Print #1, flood_name(i) & Chr(9);
        Print #1, Format(obva_n_max(i, 1), "0.000") & Chr(9);
        
        For j = 2 To 3
            Print #1, Format(obva_n_max(i, j), "0.00") & Chr(9);
        Next j
        Print #1, Format(obva_n_max(i, 4), "#0") & Chr(9);
        For j = 1 To 3
            Print #1, Format(dataf(j, i), "0.0") & Chr(9);
        Next j
        Print #1, Format(dataf(4, i), "0.000")
    Next i
    Close #1
    
    
    fn = flood_name(1)
    pn = ns(1)
    ReDim Preserve pp(pn), pqo(pn), pqs(pn), data_p(3, pn)
    For i = 1 To pn
        For j = 1 To 2
            data_p(j, i) = pqf(1, j + 4, i)
        Next j
        data_p(3, i) = q_sim_max(1, i)
    Next i
    
    Call huatu(data_p, fn)
    
    Picture1.CurrentX = 35
    Picture1.CurrentY = 106
    Picture1.Print basin & "����" & fn & "�ź�ˮ��ˮ��������"
    
    MsgBox "�������"
End Sub

'������������
Private Sub huatu(a() As Single, b As String)
    
    Dim pqmax As Single
    Dim ppmax As Single
    Dim pn As Integer '��ͼ���г���
    Dim pp() As Single '��ͼ��ˮ����
    Dim pqo() As Single '��ͼ�۲���������
    Dim pqs() As Single '��ͼģ����������
    Dim st As Integer
    Dim dc As Single, re As Single
    Dim intep As Single
    Dim inteq As Single
    
    Dim i As Integer, j As Integer
    Dim a1 As String
    Dim a2 As Single
    Dim a3 As Integer, b3 As Integer
    Dim a4 As Single, b4 As Single
    Dim a5 As Integer, b5 As Integer, c5 As Integer
    Dim a6 As Integer, b6 As Integer, c6 As Integer
    
    a3 = Val(Combo2.Text)
    a5 = Val(Combo2.List(0))
    b5 = Combo2.ListCount
    c5 = Val(Combo2.List(b5 - 2))
    pn = UBound(a, 2)
    ReDim Preserve pp(pn), pqo(pn), pqs(pn)
    
    For i = 1 To pn
        pp(i) = a(1, i)
        pqo(i) = a(2, i)
        pqs(i) = a(3, i)
    Next i
    
    pqmax = 0
    ppmax = 0
    a4 = 0
    b4 = 0
    For i = 1 To pn
        If pp(i) > ppmax Then
            ppmax = pp(i)
        End If
        If pqo(i) > pqmax Then
            pqmax = pqo(i)
        End If
        If pqs(i) > pqmax Then
            pqmax = pqs(i)
        End If
        a4 = a4 + pqo(i)
        b4 = b4 + pqs(i)
    Next i
    
    dc = nce(pqs, pqo)
    re = (b4 - a4) / a4 * 100

    Picture1.Cls
    Picture1.DrawStyle = 0

    Picture1.Scale (-5, 108)-(105, -8)
    Picture1.Line (0, 0)-(0, 100)
    Picture1.Line (0, 0)-(100, 0)
    Picture1.Line (0, 100)-(100, 100)
    Picture1.Line (100, 0)-(100, 100)
    
    Picture1.CurrentX = 100
    Picture1.CurrentY = -2
    If b = "ȫʱ��" Or b = "�����" Then
        Picture1.Print "��-��"
    Else
        Picture1.Print "Сʱ"
    End If
    
    Picture1.CurrentX = 92
    Picture1.CurrentY = 73
    Picture1.Print "��ˮ(mm)"
    Picture1.CurrentX = 1
    Picture1.CurrentY = 69
    Picture1.Print "����(m3/s)"
    Picture1.Line (75, 103)-(80, 103), RGB(255, 0, 0)
    Picture1.CurrentX = 81
    Picture1.CurrentY = 104
    Picture1.Print "�۲�����"
    Picture1.Line (90, 103)-(95, 103), RGB(0, 0, 255)
    Picture1.CurrentX = 96
    Picture1.CurrentY = 104
    Picture1.Print "ģ������"
    
    Picture1.CurrentX = -3
    Picture1.CurrentY = 104
    Picture1.Print "Ч��ϵ��: " & Format(dc, "0.000")
    
    Picture1.CurrentX = 12
    Picture1.CurrentY = 104
    Picture1.Print "�����%��: " & Format(re, "0.00")
    
    Picture1.Line (0, 70)-(100, 70)
    
    Picture1.DrawStyle = 2
    intep = (ppmax / 28 * 30) / 5
    For i = 0 To 5
        Picture1.Line (0, 100 - i * (30 / 5))-(100, 100 - i * (30 / 5))
        Picture1.CurrentX = 100.5
        Picture1.CurrentY = 100 - i * (30 / 5) + 1
        Picture1.Print Format(i * intep, "0.0")
    Next i

    inteq = (pqmax / 65 * 70) / 10
    For i = 0 To 10
        Picture1.Line (0, i * (70 / 10))-(100, i * (70 / 10))
        Picture1.CurrentX = -4
        Picture1.CurrentY = i * (70 / 10) + 1
        Picture1.Print Format(i * inteq, "#0")
    Next i
    
    If b = "�����" Then
        For i = 1 To 12
            b3 = 0
            If i = 1 Then
                b3 = 1
            Else
                For j = 1 To i - 1
                    b3 = b3 + f2(a3, j)
                Next j
                b3 = b3 + 1
            End If
            Picture1.Line (100 / (pn - 1) * (b3 - 1), 0)-(100 / (pn - 1) * (b3 - 1), 100)
            Picture1.CurrentX = 100 / (pn - 1) * (b3 - 1) - 2
            Picture1.CurrentY = -2
            Picture1.Print a3 & "-" & i
        Next i
    ElseIf b = "ȫʱ��" Then
        For i = 1 To b5 - 2
            Picture1.Line (100 / (b5 - 1) * i, 0)-(100 / (b5 - 1) * i, 100)
        Next i
        For i = 1 To (b5 - 1)
            Picture1.CurrentX = 100 / (b5 - 1) * (i - 1) - 2
            Picture1.CurrentY = -2
            Picture1.Print a5 + i - 1 & "-1"
        Next i
    Else
        a6 = Int((pn - 1) / 10)
        b6 = (pn - 1) \ a6
        
        For i = 0 To b6
            c6 = a6 * i + 1
            Picture1.Line (100 / (pn - 1) * (c6 - 1), 0)-(100 / (pn - 1) * (c6 - 1), 100)
            Picture1.CurrentX = 100 / (pn - 1) * (c6 - 1) - 1.5
            Picture1.CurrentY = -2
            Picture1.Print c6 - 1
        Next i
    End If
    
    Picture1.DrawStyle = 0
    For i = 1 To pn - 1
        Picture1.Line (100 / (pn - 1) * (i - 1), 100)-(100 / (pn - 1) * (i), 100 - pp(i) / ppmax * 28), RGB(0, 0, 255), BF
        Picture1.Line (100 / (pn - 1) * (i - 1), pqo(i) / pqmax * 65)-(100 / (pn - 1) * (i), pqo(i + 1) / pqmax * 65), RGB(255, 0, 0)
        Picture1.Line (100 / (pn - 1) * (i - 1), pqs(i) / pqmax * 65)-(100 / (pn - 1) * (i), pqs(i + 1) / pqmax * 65), RGB(0, 0, 255)
    Next i
    Picture1.Line (0, 100)-(100, 100)

End Sub

'10mm��λ������
Private Sub m36_Click()
    Randomize
    
    Dim area As Single '���������km2��
    Dim u As Single '������λת������
    
    Dim ns As Integer 'ģ��ʱ�γ�
    Dim intial(6) As Single '��ʼ������������ˮ�����ϡ��С��¡��ܣ�������ˮ��ˮ��
    Dim pe_obs() As Single '�۲⽵ˮ������
    Dim q_sim() As Single 'ģ������
    
    Dim para_in(16) As Single '������ֵ
    Dim result() As Single 'ģ����
    Dim result_day() As Single

'********************dim model parameter********************
    Dim kc As Single, um As Single, lm As Single, c As Single '����
    Dim wm As Single, b As Single, im As Single '����
    Dim sm As Single, ex As Single, kg As Single, ki As Single '��ˮԴ
    Dim ci As Single, cg As Single, cr As Single, lr As Integer '����
    
    
    Dim pqmax As Single
    Dim pn As Integer
    Dim fn As String
    Dim data_p() As Single
    
    Dim i As Long, j As Long, temp1 As Long, k As Long, h As Long
    Dim sv() As Single, ov() As Single
    Dim a2 As Single, b2 As Single, c2 As Single
    Dim a3 As Integer, b3 As Integer, c3 As Integer
    Dim a4 As String, b4 As String, c4 As String, d4 As String
    Dim a5 As Integer, b5 As Integer, c5 As Integer
    Dim a6 As Single, c6 As Integer
    Dim a7 As Integer, b7 As Integer, c7 As Integer
    Dim a8 As Single, b8 As Single
    Dim a9 As Integer, b9 As Integer
    Dim a10 As Single

'********************ȷ���������ƺ����********************
    For i = 1 To UBound(bn)
        If basin = bn(i) Then
            area = ba(i)
            Exit For
        End If
    Next i
    u = area / 3.6

'********************���ɽ�ˮ��������********************
    ns = 1000
    ReDim Preserve pe_obs(2, ns), q_sim(ns)
    pe_obs(1, 1) = 10
    pe_obs(2, 1) = 0
    For i = 2 To ns
        For j = 1 To 2
            pe_obs(j, i) = 0
        Next j
    Next i
        
    For i = 1 To 16
        para_in(i) = Val(Text1(i - 1).Text)
    Next i
    
    intial(1) = 0
    intial(2) = para_in(2)
    intial(3) = para_in(3)
    intial(4) = para_in(5) - para_in(2) - para_in(3)
    intial(5) = para_in(5)
    intial(6) = 0
    
    Call xaj_flood(para_in, pe_obs, q_sim, intial, u)
    
'������
    MSFlexGrid2.Rows = ns + 2
    MSFlexGrid2.ColWidth(0) = 1000
    MSFlexGrid2.RowHeight(0) = 500

    MSFlexGrid2.TextMatrix(0, 0) = "ʱ��" & Chr(13) & "(Hour)"
    MSFlexGrid2.TextMatrix(0, 1) = "��ˮ��" & Chr(13) & "(mm)"
    MSFlexGrid2.TextMatrix(0, 2) = "ģ������" & Chr(13) & "��m3/s��"

    For i = 0 To ns
        MSFlexGrid2.TextMatrix(i + 1, 0) = i
        MSFlexGrid2.TextMatrix(i + 1, 1) = pe_obs(1, i)
        MSFlexGrid2.TextMatrix(i + 1, 2) = Format(q_sim(i), "0.000")
    Next i

    Open App.Path & "\data\" & basin & "\ģ����\��ˮ����\10mm��λ��.txt" For Output As #1
    Print #1, "H" & Chr(9) & "P" & Chr(9) & "Q-sim"
    For k = 0 To ns
        Print #1, k & Chr(9) & pe_obs(1, k) & Chr(9) & Format(q_sim(k), "0.000")
    Next k
    Close #1
    

'��ͼ
    pn = ns
    pqmax = 0
    For i = 1 To pn
        If q_sim(i) > pqmax Then
            pqmax = q_sim(i)
        End If
    Next i

    Picture1.Cls
    Picture1.DrawStyle = 0

    Picture1.Scale (-5, 108)-(105, -8)
    Picture1.Line (0, 0)-(0, 100)
    Picture1.Line (0, 0)-(100, 0)
    Picture1.Line (0, 100)-(100, 100)
    Picture1.Line (100, 0)-(100, 100)
    
    Picture1.CurrentX = 99
    Picture1.CurrentY = -5
    Picture1.Print "Сʱ"
    
    Picture1.CurrentX = -2
    Picture1.CurrentY = 104
    Picture1.Print "����(m3/s)"
    

    inteq = (pqmax / 90 * 100) / 10
    
'    Picture1.DrawStyle = 0
'    For i = 1 To 99
'        Picture1.Line (0, i / 10 * (100 / 10))-(100, i / 10 * (100 / 10))
'    Next i
    
    Picture1.DrawStyle = 2
    For i = 0 To 10
        Picture1.Line (0, i * (100 / 10))-(100, i * (100 / 10))
        Picture1.CurrentX = -4
        Picture1.CurrentY = i * (100 / 10) + 1
        Picture1.Print Format(i * inteq, "#0")
    Next i
    
'    For i = 1 To pn - 1
'        Picture1.DrawStyle = 2
'        Picture1.Line (100 / (pn) * i, 0)-(100 / (pn) * i, 100)
'    Next i
    
    Picture1.DrawStyle = 2
    For i = 0 To pn / 10
        c6 = 10 * i
        Picture1.Line (100 / (pn) * (c6), 0)-(100 / (pn) * (c6), 100)
        Picture1.CurrentX = 100 / (pn) * (c6) - 1.5
        Picture1.CurrentY = -2
        Picture1.Print c6
    Next i
    
    Picture1.DrawStyle = 0
    For i = 0 To pn - 1
        Picture1.Line (100 / (pn) * (i), q_sim(i) / pqmax * 90)-(100 / (pn) * (i + 1), q_sim(i + 1) / pqmax * 90), RGB(0, 0, 255)
    Next

    SavePicture Picture1.Image, App.Path & "\data\" & basin & "\ģ����\��ˮ����\10mm��λ��.bmp"
    
    MsgBox "�������"
End Sub




