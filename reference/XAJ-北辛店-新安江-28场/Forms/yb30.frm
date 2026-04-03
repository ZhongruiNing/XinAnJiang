VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{008BBE7B-C096-11D0-B4E3-00A0C901D681}#1.0#0"; "TeeChart.ocx"
Begin VB.Form yb12 
   BackColor       =   &H00C0C0C0&
   Caption         =   "雨量流量过程合成图"
   ClientHeight    =   8910
   ClientLeft      =   825
   ClientTop       =   1095
   ClientWidth     =   14685
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8910
   ScaleWidth      =   14685
   WindowState     =   2  'Maximized
   Begin TeeChart.TChart TChart1 
      Height          =   7095
      Left            =   1800
      OleObjectBlob   =   "yb12.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   13215
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "yb12.frx":0905
         Height          =   1575
         Left            =   -1200
         TabIndex        =   11
         Top             =   8400
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2778
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
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
   End
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   375
      Left            =   600
      Top             =   7560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Caption         =   "Adodc8"
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   600
      Top             =   7080
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Caption         =   "Adodc7"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   495
      Left            =   1680
      Top             =   7560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Height          =   375
      Left            =   600
      Top             =   6360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   375
      Left            =   600
      Top             =   6720
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Caption         =   "Adodc6"
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   720
      Top             =   5400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Caption         =   "Adodc5"
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
   Begin VB.CommandButton Command4 
      Caption         =   "打印特征值"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打印图表"
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打印过程线"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "保存图表"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      DataSource      =   "Adodc1"
      Height          =   300
      ItemData        =   "yb12.frx":091A
      Left            =   240
      List            =   "yb12.frx":091C
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   960
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "图表缩放"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   13320
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   12240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   7440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   13920
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   13920
      Top             =   5880
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   240
      TabIndex        =   9
      Top             =   7440
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   5953
      _Version        =   393216
      DefColWidth     =   87
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
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
   Begin VB.Label Label1 
      Caption         =   "洪水选择:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.Menu mnupopup 
      Caption         =   "popup"
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
Attribute VB_Name = "yb12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1, rs2 As String
Dim rs4 As String
Dim rs5 As String
Dim rs3, rs6, rs7, rs8 As String

Dim ll As Boolean

Dim dtt, dtbg, dtnd As String
Dim dtstr As String
Private Sub Check1_Click()
If Check1.Value = 1 Then
'MsgBox "用鼠标左键由左上至右下拉框选定需放大的区域" _
       '& vbCrLf & "用鼠标右键拖动图表" _
       ' & vbCrLf & "用鼠标左键由右下至左上拉框恢复原图表"
TChart1.Zoom.Enable = True
TChart1.Scroll.Enable = pmBoth
Else: If Check1.Value = 0 Then TChart1.Zoom.Enable = False
TChart1.Scroll.Enable = pmHorizontal
End If

End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Check1.ToolTipText = "鼠标左键放大:左上至右下拉框选定区域   右键拖动图表   鼠标左键还原:右下至左上拉框"

End Sub
Private Sub Combo2_Click()

'Dim r As Integer
'Dim l As Long
'r = 0
'l = 0
'bl=true
Call drawing
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Check1.Enabled = True
End Sub

Private Sub Command1_Click()

  TChart1.Printer.ShowPreview
 
End Sub

Private Sub Command2_Click()
If ll Then
With DataReport3
Set .DataSource = Adodc5
.Sections("section4").Controls("label1").Caption = dtstr & "预报流量过程"
 .Show
End With
Else
With DataReport1
Set .DataSource = Adodc1
.Sections("section4").Controls("label1").Caption = dtstr & "预报流量过程"
 .Show
End With
End If
End Sub

Private Sub Command3_Click()
 'CommonDialog1.FileName =
 ' TChart1.Export.SaveToBitmapFile ("d:\snow.bmp")
  
  'TChart1.Export.SaveChartDialog
  TChart1.Export.ShowExport
End Sub
Private Sub Command4_Click()
Set DataReport2.DataSource = Adodc1
With DataReport2
.Sections("section2").Controls("label43").Caption = Combo2.Text
.Sections("section2").Controls("label42").Caption = "(" & dtstr & ")"
.Sections("section1").Controls("label19").Caption = Combo2.Text
.Sections("section1").Controls("label2").Caption = "(" & dtstr & ")"

On Error Resume Next

.Sections("section2").Controls("label45").Caption = Adodc6.Recordset.Fields("总雨量")
.Sections("section2").Controls("label46").Caption = Adodc6.Recordset.Fields("产流深")
.Sections("section2").Controls("label47").Caption = Adodc6.Recordset.Fields("预报水量")
.Sections("section2").Controls("label48").Caption = Adodc6.Recordset.Fields("实测水量")
.Sections("section2").Controls("label49").Caption = Adodc3.Recordset.Fields("一天水量(万立方米)")
.Sections("section2").Controls("label50").Caption = Adodc3.Recordset.Fields("五天水量(万立方米)")
.Sections("section2").Controls("label51").Caption = Adodc6.Recordset.Fields("实测洪峰")
.Sections("section2").Controls("label52").Caption = Adodc6.Recordset.Fields("实测峰现时间")
.Sections("section2").Controls("label53").Caption = Adodc6.Recordset.Fields("预报洪峰")
.Sections("section2").Controls("label54").Caption = Adodc6.Recordset.Fields("预报峰现时间")
.Sections("section2").Controls("label55").Caption = Adodc3.Recordset.Fields("总水量(万立方米)")
.Sections("section2").Controls("label56").Caption = Adodc3.Recordset.Fields("三天水量(万立方米)")
.Sections("section2").Controls("label57").Caption = Adodc3.Recordset.Fields("七天水量(万立方米)")
On Error Resume Next
.Sections("section1").Controls("label58").Caption = Adodc8.Recordset.Fields("总雨量")
.Sections("section1").Controls("label59").Caption = Adodc8.Recordset.Fields("产流深")
.Sections("section1").Controls("label60").Caption = Adodc8.Recordset.Fields("预报水量")
.Sections("section1").Controls("label61").Caption = Adodc8.Recordset.Fields("实测水量")
.Sections("section1").Controls("label62").Caption = Adodc7.Recordset.Fields("一天水量(万立方米)")
.Sections("section1").Controls("label63").Caption = Adodc7.Recordset.Fields("五天水量(万立方米)")
.Sections("section1").Controls("label64").Caption = Adodc8.Recordset.Fields("实测洪峰")
.Sections("section1").Controls("label65").Caption = Adodc8.Recordset.Fields("实测峰现时间")
.Sections("section1").Controls("label66").Caption = Adodc8.Recordset.Fields("预报洪峰")
.Sections("section1").Controls("label67").Caption = Adodc8.Recordset.Fields("预报峰现时间")
.Sections("section1").Controls("label68").Caption = Adodc7.Recordset.Fields("总水量(万立方米)")
.Sections("section1").Controls("label69").Caption = Adodc7.Recordset.Fields("三天水量(万立方米)")
.Sections("section1").Controls("label70").Caption = Adodc7.Recordset.Fields("七天水量(万立方米)")
.Show
End With
'End If
End Sub
Private Sub Form_Load()
'Combo1.Text = "断面"
'combo2.DataMember=adodc1.
' bl = False
 TChart1.Scroll.Enable = pmNone
 TChart1.Zoom.Enable = False

' Combo2.Text = "请选择洪水"
If IIS = 0 Then
 Combo2.Text = glchsdsj(1)
 rs2 = "select distinct 洪水起始时间 from ybresudp order by 洪水起始时间"
 ElseIf IIS = 1 Then
  Combo2.Text = RdDate
  rs2 = "select distinct 洪水起始时间 from ybresusdp order by 洪水起始时间"
  End If
   
 Adodc1.ConnectionString = cn
  
 Adodc2.ConnectionString = cn
 Adodc2.RecordSource = rs2
 Adodc3.ConnectionString = cn
 Adodc4.ConnectionString = cn

Adodc5.ConnectionString = cn
Adodc6.ConnectionString = cn
Adodc7.ConnectionString = cn
Adodc8.ConnectionString = cn

 
 Set Text1.DataSource = Adodc2  '////将每次洪水起始时间导入以供选择
 Adodc2.Refresh
 Text1.DataField = "洪水起始时间"
 Adodc2.Recordset.MoveFirst
 
 Dim bgdt As String
 Do While Adodc2.Recordset.EOF = False
  bgdt = Text1.Text
  Combo2.AddItem (bgdt)
  Adodc2.Recordset.MoveNext
   Do While bgdt = Text1.Text
      Adodc2.Recordset.MoveNext
   Loop
 Loop
 
 'TChart1.Axis.Bottom.Increment = TChart1.GetDateTimeStep(dtOneDay)
  
 'Combo2.Enabled = False
' Command1.Enabled = False
' Command2.Enabled = False
'Command3.Enabled = False
'Check1.Enabled = False
Call drawing


End Sub
Private Sub TChart1_OnMouseMove(ByVal Shift As TeeChart.EShiftState, ByVal X As Long, ByVal Y As Long)
'以下程序为了实现雨量和流量数值显示功能
Dim somebar, tsomebar, someline1, someline2
Dim i As Integer
Dim yin As Long

With TChart1.Series(0)
    somebar = .Clicked(X, Y)
    On Error Resume Next
        
    If somebar <> -1 Then
    tsomebar = somebar
    '.PointColor(somebar) = vbBlue
     
     Adodc1.Recordset.MoveFirst
     For i = 1 To somebar
     If Not Adodc1.Recordset.EOF Then
          Adodc1.Recordset.MoveNext
          End If
     Next i
     
     yin = Adodc1.Recordset.Fields("dt")
     
        
      TChart1.ToolTipText = todate(yin) _
      + "  " + Format(.YValues.Value(somebar), "##.##") + " mm"
    Else
  
    End If
End With

With TChart1.Series(1)
    someline1 = .Clicked(X, Y)
    On Error Resume Next
    
    If someline1 <> -1 Then
    
    Adodc1.Recordset.MoveFirst
    For i = 1 To someline1
    If Not Adodc1.Recordset.EOF Then
     Adodc1.Recordset.MoveNext
     End If
     Next i
     yin = Adodc1.Recordset.Fields("dt")
 
  ' Text3.Text = .XValues.Value(someline1)
     ' TChart1.ToolTipText = Str(.XValues.Value(someline1)) _
      '+ "  " + Str(.YValues.Value(someline1)) + " m3"
      TChart1.ToolTipText = todate(yin) _
      + "  " + Format(.YValues.Value(someline1), "####.##") + " m3"

    Else
    End If
   
End With


With TChart1.Series(2)
    someline2 = .Clicked(X, Y)
    On Error Resume Next
    
    If someline2 <> -1 Then
    
    Adodc1.Recordset.MoveFirst
    For i = 1 To someline2
     If Not Adodc1.Recordset.EOF Then
     Adodc1.Recordset.MoveNext
     End If
     Next i
     yin = Adodc1.Recordset.Fields("dt")
     
      'TChart1.ToolTipText = Format(.XValues.Value(someline2), "yyyy-mm-dd hh:nn:ss") _
     ' + "  " + Format(.YValues.Value(someline2), "0.00") + " m3"
     TChart1.ToolTipText = todate(yin) _
      + "  " + Format(.YValues.Value(someline2), "#####.##") + " m3"

    Else
    End If
End With


With TChart1.Series(3)
    someline2 = .Clicked(X, Y)
    On Error Resume Next
    
    If someline2 <> -1 Then
    
    Adodc4.Recordset.MoveFirst
    For i = 1 To someline2
     If Not Adodc4.Recordset.EOF Then
     Adodc4.Recordset.MoveNext
     End If
     Next i
     yin = Adodc4.Recordset.Fields("dt")
     
      'TChart1.ToolTipText = Format(.XValues.Value(someline2), "yyyy-mm-dd hh:nn:ss") _
     ' + "  " + Format(.YValues.Value(someline2), "0.00") + " m3"
     TChart1.ToolTipText = todate(yin) _
      + "  " + Format(.YValues.Value(someline2), "#####.##") + " m3"

    Else
    End If
End With
End Sub

'TChart1.ToolTipText = ""
Function todate(Dt As Long)
'将日期从长整形转化为普通的形式

Dim Y, m, d, t As Integer
 Y = Int(Dt / 1000000)
m = Int((Dt Mod 1000000) / 10000)
d = Int(((Dt Mod 1000000) Mod 10000) / 100)
t = ((Dt Mod 1000000) Mod 10000) Mod 100
 todate = Y & "年" & m & "月" & d & "日" & t & "时"
 
End Function

Sub drawing()
TChart1.Zoom.Undo
TChart1.Scroll.Enable = pmHorizontal

dtt = Combo2.Text
If IIS = 0 Then
rs5 = "select 洪水起始时间 from jyresudp where  洪水起始时间=" & dtt

ElseIf IIS = 1 Then
rs5 = "select 洪水起始时间 from jyresusdp where  洪水起始时间=" & dtt

End If
Adodc5.RecordSource = rs5
Adodc5.Refresh
If Adodc5.Recordset.EOF And Adodc5.Recordset.BOF Then
ll = False
Else
ll = True
End If

If IIS = 0 Then
 rs1 = "select * From  ybresudp  where 洪水起始时间=" & dtt & " order by dt"
rs4 = "select * From  jyresudp  where 洪水起始时间=" & dtt & " order by dt"
rs5 = "select ybresudp.dt, ybresudp.面平均雨量,ybresudp.实测流量,ybresudp.预报流量,jyresudp.预报流量 as jy预报流量  From  ybresudp,jyresudp  where ybresudp.洪水起始时间=" & dtt & " and jyresudp.dt=ybresudp.dt"
rs3 = "select * from  ybresudp0  where  洪水起始时间=" & dtt
rs6 = "select * from  ybresudp1 where 洪水起始时间=" & dtt
rs7 = "select * from  jyresudp0 where 洪水起始时间=" & dtt
rs8 = "select * from  jyresudp1 where 洪水起始时间=" & dtt
ElseIf IIS = 1 Then
rs1 = "select * From  ybresusdp  where 洪水起始时间=" & dtt & " order by dt"
rs4 = "select * From  jyresusdp  where 洪水起始时间=" & dtt & " order by dt"
rs5 = "select ybresusdp.dt, ybresusdp.面平均雨量,ybresusdp.实测流量,ybresusdp.预报流量,jyresusdp.预报流量 as jy预报流量  From  ybresusdp,jyresusdp  where ybresusdp.洪水起始时间=" & dtt & " and jyresusdp.dt=ybresusdp.dt"
rs3 = "select * from  ybresusdp0 where 洪水起始时间=" & dtt
rs6 = "select * from  ybresusdp1 where 洪水起始时间=" & dtt
rs7 = "select * from  jyresusdp0 where 洪水起始时间=" & dtt
rs8 = "select * from  jyresusdp1 where 洪水起始时间=" & dtt
End If


Adodc1.RecordSource = rs1
Adodc4.RecordSource = rs4
Adodc5.RecordSource = rs5
Adodc3.RecordSource = rs3
Adodc6.RecordSource = rs6
Adodc7.RecordSource = rs7
Adodc8.RecordSource = rs8

Adodc1.Refresh
Adodc4.Refresh
Adodc5.Refresh
Adodc3.Refresh
Adodc6.Refresh
Adodc7.Refresh
Adodc8.Refresh

Set DataGrid1.DataSource = Adodc1

TChart1.Series(0).Clear
TChart1.Series(1).Clear
TChart1.Series(2).Clear
TChart1.Series(3).Clear

With Adodc1.Recordset
If .EOF And .BOF Then
MsgBox "此次洪水数据不存在"
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Check1.Enabled = False
Exit Sub
End If
.MoveFirst
dtbg = todate(.Fields("dt"))
 While Not .EOF
 dtt = todate(.Fields("dt"))
  
  TChart1.Series(0).Add .Fields("面平均雨量"), dtt, clTeeColor
  TChart1.Series(1).Add .Fields("实测流量"), dtt, clTeeColor
  TChart1.Series(2).Add .Fields("预报流量"), dtt, clTeeColor
  
'errhandle:
'    MsgBox "数据库访问出错"
On Error Resume Next
 .MoveNext
  Wend
  dtnd = dtt
End With

With Adodc4.Recordset

If Not .EOF Or Not .BOF Then
.MoveFirst

While Not .EOF

dtt = todate(.Fields("ybresudp.dt"))
TChart1.Series(3).Add .Fields("预报流量"), dtt, clTeeColor

On Error Resume Next
.MoveNext
Wend
End If
End With



dtstr = dtbg & "_" & dtnd

TChart1.Header.Text.Clear
TChart1.Header.Text.Add ("预报流量过程线")
TChart1.Header.Text.Add ("(" & dtstr & ")")


 'Command1.Enabled = True
' Command2.Enabled = True
'Command3.Enabled = True
'Check1.Enabled = True
End Sub











