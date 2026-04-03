VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form yb01 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选取要验证的年份"
   ClientHeight    =   7725
   ClientLeft      =   795
   ClientTop       =   1065
   ClientWidth     =   11850
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   11850
   StartUpPosition =   1  '所有者中心
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   315
      Left            =   5040
      Top             =   600
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   315
      Left            =   3360
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   2040
      Top             =   480
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   720
      Top             =   480
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cl 
      Caption         =   "取消"
      Height          =   375
      Left            =   10080
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton ok 
      Caption         =   "确定"
      Height          =   375
      Left            =   8640
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   6135
      Left            =   8040
      TabIndex        =   4
      Top             =   720
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   10821
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "选择要率定的年份"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6615
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   11668
      _Version        =   393216
      DefColWidth     =   47
      HeadLines       =   1
      RowHeight       =   20
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "现有洪水年份"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Cmddel1 
      Caption         =   "删除"
      Height          =   375
      Index           =   1
      Left            =   9240
      TabIndex        =   2
      Top             =   7080
      Width           =   855
   End
   Begin VB.CommandButton Cmdadd1 
      Caption         =   "增加"
      Height          =   375
      Left            =   8160
      TabIndex        =   1
      Top             =   7080
      Width           =   855
   End
   Begin VB.CommandButton Cmddel0 
      Caption         =   "删除"
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   0
      Top             =   7080
      Width           =   975
   End
End
Attribute VB_Name = "yb01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cl_Click()
  Unload yb01
End Sub
Private Sub Cmdadd0_Click()
Dim i As Integer
On Error GoTo c
i = 0
If Not Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveFirst
End If
If Not Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveLast
i = Adodc1.Recordset(0)
Else
i = 0
End If
Adodc1.Recordset.AddNew
Adodc1.Recordset(0) = i + 1
Adodc1.Recordset(1) = 0
Adodc1.Recordset.Update
If Not Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveFirst
End If
Adodc1.Recordset.MoveLast
c: Exit Sub
End Sub
Private Sub Cmdadd1_Click()
Dim i As Integer, nh As Integer, hh As Long, rhh() As Single
On Error GoTo c
i = 0
If Not Adodc3.Recordset.BOF Then
Adodc3.Recordset.MoveFirst
End If
If Not Adodc3.Recordset.EOF Then
Adodc3.Recordset.MoveLast
i = Adodc3.Recordset(0)
Else
i = 0
End If
Adodc3.Recordset.AddNew
Adodc3.Recordset(0) = i + 1
Adodc3.Recordset.Update
If Not Adodc3.Recordset.BOF Then
Adodc3.Recordset.MoveFirst
End If
Adodc3.Recordset.MoveLast
c: Exit Sub
End Sub
Private Sub Cmddel0_Click(Index As Integer)
On Error GoTo c
Adodc1.Recordset.Delete
If Not Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveFirst
End If
Adodc1.Recordset.MoveLast
c: Exit Sub
End Sub
Private Sub Cmddel1_Click(Index As Integer)
Dim i As Integer, nh As Integer, hh() As Long
On Error GoTo c
Adodc3.Recordset.Delete
If Not Adodc3.Recordset.BOF Then
Adodc3.Recordset.MoveFirst
End If
Adodc3.Recordset.MoveLast
c: Exit Sub
End Sub

Private Sub Form_Load()


bname = "dayflood" + CStr(dyly)

sql1 = "select * from " + bname + "  order by [No]"

 With Adodc1
    .ConnectionString = kname
    .RecordSource = bname
  End With
  Set DataGrid1.DataSource = Adodc1
 
 DataGrid1.Columns(0).Width = 600
 DataGrid1.Columns(1).Width = 1000
 DataGrid1.Columns(2).Width = 2400
 DataGrid1.Columns(3).Width = 2400

bname = "selecyear" + Basin
With Adodc3
  .ConnectionString = kname
  .RecordSource = bname
End With
Set DataGrid3.DataSource = Adodc3


End Sub

Private Sub ok_Click()
  Unload Me
End Sub
