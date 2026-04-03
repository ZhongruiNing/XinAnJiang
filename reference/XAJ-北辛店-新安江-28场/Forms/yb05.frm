VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form yb05 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选取要验证的洪水"
   ClientHeight    =   7305
   ClientLeft      =   795
   ClientTop       =   1065
   ClientWidth     =   14340
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   14340
   StartUpPosition =   1  '所有者中心
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   312
      Left            =   5280
      Top             =   240
      Visible         =   0   'False
      Width           =   972
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
      Height          =   312
      Left            =   3600
      Top             =   240
      Visible         =   0   'False
      Width           =   1332
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
      RecordSource    =   "totalbasin"
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
      Height          =   372
      Left            =   2160
      Top             =   120
      Visible         =   0   'False
      Width           =   972
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
      Height          =   372
      Left            =   840
      Top             =   120
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
      Left            =   13200
      TabIndex        =   8
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton ok 
      Caption         =   "确定"
      Height          =   375
      Left            =   11880
      TabIndex        =   7
      Top             =   0
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "yb05.frx":0000
      Height          =   1095
      Left            =   11280
      TabIndex        =   6
      Top             =   4920
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1931
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   25
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
      Caption         =   "选择要率定的流域"
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   5775
      Left            =   9360
      TabIndex        =   5
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   10186
      _Version        =   393216
      DefColWidth     =   67
      HeadLines       =   1
      RowHeight       =   25
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
      Caption         =   "选定率定的洪水"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   3975
      Left            =   11280
      TabIndex        =   4
      Top             =   720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   7011
      _Version        =   393216
      DefColWidth     =   67
      HeadLines       =   1
      RowHeight       =   25
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
      Caption         =   "需要率定的流域及代码"
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
      Height          =   5895
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   10398
      _Version        =   393216
      DefColWidth     =   142
      HeadLines       =   1
      RowHeight       =   25
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
      Caption         =   "现有洪水场次"
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
      Left            =   10320
      TabIndex        =   2
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton Cmdadd1 
      Caption         =   "增加"
      Height          =   375
      Left            =   9360
      TabIndex        =   1
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton Cmddel0 
      Caption         =   "删除"
      Height          =   375
      Index           =   0
      Left            =   6840
      TabIndex        =   0
      Top             =   6480
      Width           =   975
   End
End
Attribute VB_Name = "yb05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cl_Click()
  Unload yb05
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
yb05.Caption = "要验证和率定的洪水资料及流域选取"

bname = "dataflood" + dyly
sql1 = "select * from " + bname + " order by No  "

With Adodc1
  .ConnectionString = kname
  .RecordSource = bname
End With
Set DataGrid1.DataSource = Adodc1

 'DataGrid1.Columns(0).Width = 600
 'DataGrid1.Columns(1).Width = 1000
 'DataGrid1.Columns(2).Width = 2400
 'DataGrid1.Columns(3).Width = 2400

bname = "totalbasin"
With Adodc2
  .ConnectionString = kname
  .RecordSource = bname
End With
Set DataGrid2.DataSource = Adodc2

bname = "selecflood" + dyly
With Adodc3
  .ConnectionString = kname
  .RecordSource = bname
End With
Set DataGrid3.DataSource = Adodc3

bname = "selecbasin"
With Adodc4
  .ConnectionString = kname
  .RecordSource = bname
End With
Set DataGrid4.DataSource = Adodc4
End Sub

Private Sub ok_Click()
  Unload yb05
  Call findbasin
 End Sub
