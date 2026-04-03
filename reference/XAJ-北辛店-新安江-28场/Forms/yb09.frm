VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{B6C10482-FB89-11D4-93C9-006008A7EED4}#1.0#0"; "TeeChart5.ocx"
Begin VB.Form yb09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "雨量流量过程合成图"
   ClientHeight    =   7428
   ClientLeft      =   816
   ClientTop       =   1080
   ClientWidth     =   12444
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   15264
   WindowState     =   2  'Maximized
   Begin TeeChart.TChart TChart1 
      Height          =   3132
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   13572
      Base64          =   $"yb09.frx":0000
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   9840
      Top             =   9000
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   593
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6132
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   14928
      _ExtentX        =   26331
      _ExtentY        =   10816
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "流量过程预报"
      TabPicture(0)   =   "yb09.frx":0C06
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "预报特征值"
      TabPicture(1)   =   "yb09.frx":0C22
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "DataGrid2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "时段特征值"
      TabPicture(2)   =   "yb09.frx":0C3E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DataGrid3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   5292
         Left            =   -75000
         TabIndex        =   13
         Top             =   360
         Width           =   14772
         _ExtentX        =   26056
         _ExtentY        =   9335
         _Version        =   393216
         DefColWidth     =   133
         HeadLines       =   1
         RowHeight       =   27
         RowDividerStyle =   3
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   5292
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   14532
         _ExtentX        =   25633
         _ExtentY        =   9335
         _Version        =   393216
         DefColWidth     =   83
         HeadLines       =   1
         RowHeight       =   25
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5172
         Left            =   -74880
         TabIndex        =   11
         Top             =   360
         Width           =   14532
         _ExtentX        =   25633
         _ExtentY        =   9123
         _Version        =   393216
         DefColWidth     =   93
         HeadLines       =   1
         RowHeight       =   27
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   5880
      Top             =   8400
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   593
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
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   12000
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   7320
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   13080
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "图表缩放"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      DataSource      =   "Adodc1"
      Height          =   276
      ItemData        =   "yb09.frx":0C5A
      Left            =   0
      List            =   "yb09.frx":0C5C
      TabIndex        =   5
      Text            =   "Combo2"
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "保存图表"
      Height          =   492
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打印过程线"
      Height          =   492
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打印图表"
      Height          =   492
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "打印特征值"
      Height          =   492
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   3960
      Top             =   8280
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   656
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   13680
      Top             =   5160
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3408
      _ExtentY        =   593
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
      Left            =   13680
      Top             =   5760
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   593
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
   Begin VB.Label Label1 
      Caption         =   "年份选择:"
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   1092
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
Attribute VB_Name = "yb09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1, rs2, rs3 As String

Dim dtt, dtbg, dtnd As String
Dim dtt1 As Date
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

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Check1.ToolTipText = "鼠标左键放大:左上至右下拉框选定区域   右键拖动图表   鼠标左键还原:右下至左上拉框"

End Sub
Private Sub Combo1_Click()
'Combo2.Enabled = True
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
With DataReport1
Set .DataSource = Adodc1
 .Sections("section4").Controls("label1").Caption = dtstr & "预报流量过程"
 .Show
End With
'TChart1.Series(2).LabelsSource = "实测流量"
End Sub

Private Sub Command3_Click()
 'CommonDialog1.FileName =
 ' TChart1.Export.SaveToBitmapFile ("d:\snow.bmp")
  
  'TChart1.Export.SaveChartDialog
  TChart1.Export.ShowExport
End Sub


Private Sub Command4_Click()

With DataReport2
Set .DataSource = Adodc3

.Sections("section1").Controls("label19").Caption = Combo2.Text
.Sections("section1").Controls("label2").Caption = "(" & dtstr & ")"

.Show
End With
End Sub



Private Sub Form_Load()
Dim it As Long, yy As Integer, dd As Integer, mm As Integer, hh As Integer
'Combo1.Text = "断面"
'combo2.DataMember=adodc1.
' bl = False
 TChart1.Scroll.Enable = pmNone
 TChart1.Zoom.Enable = False

' Combo2.Text = "请选择洪水"
 it = glchsdsj(1)
 Call ymdh(it, yy, mm, dd, hh)
 Combo2.Text = it
  rs2 = "select 洪水起始时间  from" + " ybresu" + CStr(dyly) + " order by 洪水起始时间 "
   
 
 Adodc1.ConnectionString = cn
 
 Adodc2.ConnectionString = cn
 Adodc2.RecordSource = rs2
 
 Adodc3.ConnectionString = cn


 
 Set Text1.DataSource = Adodc2  '////将每次洪水起始时间导入以供选择

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
 
Call drawing
 
bname = "ybresu" + CStr(dyly)
sql1 = "select *  from " + bname + " order by 洪水起始时间"

With Adodc3
  .ConnectionString = kname
  .RecordSource = "select *  from " + bname + " order by dt"
End With
Set DataGrid1.DataSource = Adodc3

bname = "ybresu" + CStr(dyly) + "1"
sql1 = "select * from " + bname + " order by 洪水起始时间 "
With Adodc4
  .ConnectionString = kname
  .RecordSource = sql1
End With
Set DataGrid2.DataSource = Adodc4

bname = "ybresu" + CStr(dyly) + "1"
sql1 = "select * from " + bname + " order by 洪水起始时间 "
With Adodc5
  .ConnectionString = kname
  .RecordSource = sql1
End With
Set DataGrid3.DataSource = Adodc5

Adodc3.Visible = False
Adodc4.Visible = False
Adodc5.Visible = False
 

End Sub

Private Sub TChart1_OnMouseMove(ByVal Shift As TeeChart.EShiftState, ByVal x As Long, ByVal Y As Long)
'以下程序为了实现雨量和流量数值显示功能
Dim somebar, tsomebar, someline1, someline2
Dim i As Integer
Dim yin As Long

With TChart1.Series(0)
    somebar = .Clicked(x, Y)
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
    someline1 = .Clicked(x, Y)
    If someline1 <> -1 Then
    
    Adodc1.Recordset.MoveFirst
    For i = 1 To someline1
    If Not Adodc1.Recordset.EOF Then
     Adodc1.Recordset.MoveNext
     End If
     Next i
     yin = Adodc1.Recordset.Fields("dt")
 
   'Text3.Text = .XValues.Value(someline1)
    'TChart1.ToolTipText = Str(.XValues.Value(someline1)) + "  " + Str(.YValues.Value(someline1)) + " m3"
     TChart1.ToolTipText = todate(yin) + "  " + Format(.YValues.Value(someline1), "####.##") + " m3/s"

    Else
    End If
   
End With


With TChart1.Series(2)
    someline2 = .Clicked(x, Y)
    If someline2 <> -1 Then
    
    Adodc1.Recordset.MoveFirst
    For i = 1 To someline2
     If Not Adodc1.Recordset.EOF Then
     Adodc1.Recordset.MoveNext
     End If
     Next i
     yin = Adodc1.Recordset.Fields("dt")
     
     'TChart1.ToolTipText = Format(.XValues.Value(someline2), "yyyy-mm-dd hh:nn:ss") + "  " + Format(.YValues.Value(someline2), "0.00") + " m3"
     TChart1.ToolTipText = todate(yin) + "  " + Format(.YValues.Value(someline2), "#####.##") + " m3"

    Else
    End If
End With


With TChart1.Series(3)
    someline3 = .Clicked(x, Y)
    If someline3 <> -1 Then

    Adodc1.Recordset.MoveFirst
    For i = 1 To someline3
     If Not Adodc1.Recordset.EOF Then
     Adodc1.Recordset.MoveNext
     End If
     Next i
     yin = Adodc1.Recordset.Fields("dt")

     'TChart1.ToolTipText = Format(.XValues.Value(someline2), "yyyy-mm-dd hh:nn:ss") + "  " + Format(.YValues.Value(someline2), "0.00") + " m3"
     TChart1.ToolTipText = todate(yin) _
      + "  " + Format(.YValues.Value(someline3), "#####.##") + " m3" '''

    Else
    End If
End With

End Sub

'TChart1.ToolTipText = ""
Function todate(dt As Long)
'将日期从长整形转化为普通的形式

Dim Y, m, d, t As Integer
Dim yy As Integer, h As Integer
Dim y1 As String


yy = Int(dt / 1000000)
m = Int((dt Mod 1000000) / 100)
d = Int(((dt Mod 1000000) Mod 100) / 1)
t = ((dt Mod 100000000) Mod 10000) Mod 100
t = ((dt Mod 100000000) Mod 10000) Mod 100
y1 = Mid(CStr(yy), 3, 4)
Y = Val(y1)

todate = Y & "." & m & "." & d & "."
 
End Function

Sub drawing()
TChart1.Zoom.Undo
TChart1.Scroll.Enable = pmHorizontal

dtt = Combo2.Text
rs1 = "select * From  " + "ybresu" + CStr(dyly) + "  where 洪水起始时间 =" & dtt & " order by dt"

Adodc1.RecordSource = rs1

Adodc1.Refresh

TChart1.Series(0).Clear
TChart1.Series(1).Clear
TChart1.Series(2).Clear
TChart1.Series(3).Clear

With Adodc1.Recordset
If .EOF And .BOF Then
MsgBox "此年份不存在"
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
 
  dtt1 = .Fields("dt1")
  TChart1.Series(0).XValues.DateTime = True
  TChart1.Series(1).XValues.DateTime = True
  TChart1.Series(2).XValues.DateTime = True
  TChart1.Series(3).XValues.DateTime = True
  
  
  
  TChart1.Series(0).Add .Fields("面平均雨量"), dtt1, clTeeColor
  TChart1.Series(1).Add .Fields("实测流量"), dtt1, clTeeColor
  TChart1.Series(2).Add .Fields("预报流量"), dtt1, clTeeColor
  TChart1.Series(3).Add .Fields("上游来水演算流量"), dtt1, clTeeColor
'errhandle:
'    MsgBox "数据库访问出错"
On Error Resume Next

 .MoveNext
 
  Wend
  dtnd = dtt
  
End With
dtstr = dtbg & " ~ " & dtnd
TChart1.Header.Text.Clear
TChart1.Header.Text.Add (dylyc + "站预报流量过程线")
TChart1.Header.Text.Add ("(" & dtstr & ")")



End Sub









