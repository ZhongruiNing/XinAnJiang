VERSION 5.00
Object = "{9D6ED199-5910-11D2-98A6-00A0C9742CCA}#4.0#0"; "mapx40.ocx"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   11220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13695
   LinkTopic       =   "Form4"
   ScaleHeight     =   11220
   ScaleWidth      =   13695
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   7680
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   4080
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin MapXLib.Map Map1 
      Height          =   9015
      Left            =   1440
      TabIndex        =   0
      Top             =   1560
      Width           =   10815
      _Version        =   400011
      _ExtentX        =   19076
      _ExtentY        =   15901
      _StockProps     =   1
      GeoDictionary   =   "GeoDictionary"
      GeoSet          =   "Asia"
      GeoSetUserName  =   "Asia"
      CurrentTool     =   1000
      Zoom            =   10480
      MaxSearchTime   =   30
      CenterX         =   106.509527
      CenteryY        =   32.939751
      Rotation        =   0
      TitleText       =   "Asia"
      DataSetGeoField =   ""
      AutoRedraw      =   -1  'True
      PreferCompactLegends=   0   'False
      TitleVisible    =   -1  'True
      MousePointer    =   0
      MouseIcon       =   ""
      MatchThreshold  =   80
      WaitCursorEnabled=   -1  'True
      MousewheelSupport=   1
      MatchNumericFields=   0   'False
      RedrawInterval  =   10
      PanAnimationLayer=   0   'False
      InfotipSupport  =   -1  'True
      InfotipPopupDelay=   500
      DefaultConversionResolution=   12
      ExportSelection =   0   'False
      NumLayers       =   5
      Layer0.path     =   "Asiacaps.TAB"
      Layer0.name     =   "Asia Capitals"
      Layer0.visible  =   -1  'True
      Layer0.selectable=   -1  'True
      Layer0.editable =   0   'False
      Layer0.shownodes=   0   'False
      Layer0.showcentroids=   0   'False
      Layer0.showlinedirection=   0   'False
      Layer0.autolabel=   -1  'True
      Layer0.zoomlayering=   -1  'True
      Layer0.minzoom  =   3000
      Layer0.maxzoom  =   10000
      Layer0.DrawLabelsAfter=   0   'False
      Layer0.styleoverride=   0   'False
      Layer0.labelstyle.TextFontBackColor=   13696976
      Layer0.labelstyle.TextFontHalo=   -1  'True
      Layer0.labelstyle.SymbolChar=   0
      BeginProperty Layer0.labelstyle.TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Layer0.labelstyle.LineStyle=   1
      Layer0.labelstyle.LineWidth=   1
      Layer0.LabelMax =   100
      Layer0.LabelOverlap=   0   'False
      Layer0.LabelDuplicate=   0   'False
      Layer0.LabelOffset=   2
      Layer0.LabelLineType=   2
      Layer0.LabelZoomMax=   10000
      Layer0.LabelZoomMin=   0
      Layer0.LabelZoom=   0   'False
      Layer0.LabelVisible=   -1  'True
      Layer0.LabelOrientation=   5
      Layer0.LabelParellel=   -1  'True
      Layer1.path     =   "Asicty79.TAB"
      Layer1.name     =   "Asia Major Cities"
      Layer1.visible  =   -1  'True
      Layer1.selectable=   -1  'True
      Layer1.editable =   0   'False
      Layer1.shownodes=   0   'False
      Layer1.showcentroids=   0   'False
      Layer1.showlinedirection=   0   'False
      Layer1.autolabel=   -1  'True
      Layer1.zoomlayering=   -1  'True
      Layer1.minzoom  =   0
      Layer1.maxzoom  =   3000
      Layer1.DrawLabelsAfter=   0   'False
      Layer1.styleoverride=   0   'False
      Layer1.labelstyle.TextFontBackColor=   13696976
      Layer1.labelstyle.TextFontHalo=   -1  'True
      Layer1.labelstyle.SymbolChar=   0
      BeginProperty Layer1.labelstyle.TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Layer1.labelstyle.LineStyle=   1
      Layer1.labelstyle.LineWidth=   1
      Layer1.LabelMax =   100
      Layer1.LabelOverlap=   0   'False
      Layer1.LabelDuplicate=   0   'False
      Layer1.LabelOffset=   2
      Layer1.LabelLineType=   2
      Layer1.LabelZoomMax=   10000
      Layer1.LabelZoomMin=   0
      Layer1.LabelZoom=   0   'False
      Layer1.LabelVisible=   -1  'True
      Layer1.LabelOrientation=   5
      Layer1.LabelParellel=   -1  'True
      Layer2.path     =   "Europe.TAB"
      Layer2.name     =   "Europe"
      Layer2.visible  =   -1  'True
      Layer2.selectable=   -1  'True
      Layer2.editable =   0   'False
      Layer2.shownodes=   0   'False
      Layer2.showcentroids=   0   'False
      Layer2.showlinedirection=   0   'False
      Layer2.autolabel=   0   'False
      Layer2.zoomlayering=   0   'False
      Layer2.minzoom  =   0
      Layer2.maxzoom  =   0
      Layer2.DrawLabelsAfter=   0   'False
      Layer2.styleoverride=   -1  'True
      Layer2.layerstyle.TextFontBackColor=   16777215
      Layer2.layerstyle.SymbolType=   2
      Layer2.layerstyle.SupportsBitmapSymbols=   -1  'True
      Layer2.layerstyle.SymbolVectorSize=   12
      BeginProperty Layer2.layerstyle.TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Layer2.layerstyle.SymbolFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Map Symbols"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Layer2.layerstyle.LineStyle=   1
      Layer2.layerstyle.LineWidth=   1
      Layer2.layerstyle.RegionColor=   13684944
      Layer2.layerstyle.LinePattern=   2
      Layer2.layerstyle.RegionBackColor=   16777215
      Layer2.layerstyle.RegionBorderStyle=   1
      Layer2.layerstyle.RegionBorderWidth=   1
      Layer2.labelstyle.TextFontBackColor=   16777215
      Layer2.labelstyle.SymbolChar=   0
      BeginProperty Layer2.labelstyle.TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Layer2.labelstyle.LineStyle=   1
      Layer2.labelstyle.LineWidth=   1
      Layer2.LabelMax =   100
      Layer2.LabelOverlap=   0   'False
      Layer2.LabelDuplicate=   -1  'True
      Layer2.LabelOffset=   2
      Layer2.LabelLineType=   0
      Layer2.LabelZoomMax=   10000
      Layer2.LabelZoomMin=   0
      Layer2.LabelZoom=   0   'False
      Layer2.LabelVisible=   -1  'True
      Layer2.LabelOrientation=   0
      Layer2.LabelParellel=   -1  'True
      Layer3.path     =   "Asia.TAB"
      Layer3.name     =   "Asia"
      Layer3.visible  =   -1  'True
      Layer3.selectable=   -1  'True
      Layer3.editable =   0   'False
      Layer3.shownodes=   0   'False
      Layer3.showcentroids=   0   'False
      Layer3.showlinedirection=   0   'False
      Layer3.autolabel=   -1  'True
      Layer3.zoomlayering=   0   'False
      Layer3.minzoom  =   0
      Layer3.maxzoom  =   0
      Layer3.DrawLabelsAfter=   0   'False
      Layer3.styleoverride=   0   'False
      Layer3.labelstyle.TextFontColor=   128
      Layer3.labelstyle.TextFontBackColor=   13696976
      Layer3.labelstyle.TextFontHalo=   -1  'True
      Layer3.labelstyle.SymbolChar=   0
      BeginProperty Layer3.labelstyle.TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Layer3.labelstyle.LineStyle=   1
      Layer3.labelstyle.LineWidth=   1
      Layer3.LabelMax =   100
      Layer3.LabelOverlap=   0   'False
      Layer3.LabelDuplicate=   0   'False
      Layer3.LabelOffset=   2
      Layer3.LabelLineType=   0
      Layer3.LabelZoomMax=   10000
      Layer3.LabelZoomMin=   0
      Layer3.LabelZoom=   0   'False
      Layer3.LabelVisible=   -1  'True
      Layer3.LabelOrientation=   0
      Layer3.LabelParellel=   -1  'True
      Layer4.path     =   "Ocn_asia.TAB"
      Layer4.name     =   "Ocean (for Asia Maps)"
      Layer4.visible  =   -1  'True
      Layer4.selectable=   0   'False
      Layer4.editable =   0   'False
      Layer4.shownodes=   0   'False
      Layer4.showcentroids=   0   'False
      Layer4.showlinedirection=   0   'False
      Layer4.autolabel=   0   'False
      Layer4.zoomlayering=   0   'False
      Layer4.minzoom  =   0
      Layer4.maxzoom  =   0
      Layer4.DrawLabelsAfter=   0   'False
      Layer4.styleoverride=   0   'False
      Layer4.labelstyle.TextFontBackColor=   16777215
      Layer4.labelstyle.SymbolChar=   0
      BeginProperty Layer4.labelstyle.TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Layer4.labelstyle.LineStyle=   1
      Layer4.labelstyle.LineWidth=   1
      Layer4.LabelMax =   100
      Layer4.LabelOverlap=   0   'False
      Layer4.LabelDuplicate=   -1  'True
      Layer4.LabelOffset=   2
      Layer4.LabelLineType=   0
      Layer4.LabelZoomMax=   10000
      Layer4.LabelZoomMin=   0
      Layer4.LabelZoom=   0   'False
      Layer4.LabelVisible=   -1  'True
      Layer4.LabelOrientation=   0
      Layer4.LabelParellel=   -1  'True
      TitleStyle.TextFontBackColor=   16777215
      TitleStyle.TextFontOpaque=   -1  'True
      TitleStyle.SymbolChar=   0
      BeginProperty TitleStyle.TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   45
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleStyle.SymbolFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   45
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DefaultStyle.TextFontBackColor=   13430215
      DefaultStyle.SupportsBitmapSymbols=   -1  'True
      DefaultStyle.SymbolChar=   55
      DefaultStyle.SymbolFontBackColor=   13430215
      BeginProperty DefaultStyle.TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DefaultStyle.SymbolFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Map Symbols"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DefaultStyle.LineStyle=   1
      DefaultStyle.LineWidth=   1
      DefaultStyle.RegionColor=   16777215
      DefaultStyle.LinePattern=   2
      DefaultStyle.RegionBackColor=   13430215
      DefaultStyle.RegionBorderStyle=   1
      DefaultStyle.RegionBorderWidth=   1
      HasProjectionInfo=   -1  'True
      NumericCoordsys =   "Form4.frx":0000
      DisplayCoordsys =   "Form4.frx":0130
      NumDatasets     =   0
      TitleX          =   5000
      TitleY          =   1000
      TitleVisible    =   -1  'True
      TitleEditable   =   -1  'True
      TitlePostiion   =   0
      TitleBorder     =   -1  'True
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Map1.CurrentTool = miSelectTool
    
End Sub

Private Sub Command2_Click()
    Map1.CurrentTool = miZoomInTool

End Sub

Private Sub Command3_Click()
    Map1.CurrentTool = miZoomOutTool
End Sub

Private Sub Command4_Click()
    Map1.CurrentTool = miPanTool
End Sub

Private Sub Command5_Click()
    Map1.Layers.LayersDlg
End Sub

