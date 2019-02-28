VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form frmSharpGrid 
   Caption         =   "#GridForm"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   9165
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   19595265
      CurrentDate     =   38306
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   600
      Width           =   2655
   End
   Begin VB.PictureBox ctlEBCalendar1 
      Height          =   315
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      Caption         =   "Dil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000009&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   2415
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewer21 
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5318
      SectionData     =   "frmSharpGrid.frx":0000
   End
   Begin DDSharpGridOLEDB2.SGGrid SGGrid1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7815
      _cx             =   13785
      _cy             =   8705
      DataMember      =   ""
      DataMode        =   1
      AutoFields      =   -1  'True
      Enabled         =   -1  'True
      GridBorderStyle =   1
      ScrollBars      =   3
      FlatScrollBars  =   0
      ScrollBarTrack  =   0   'False
      DataRowCount    =   2
      BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataColCount    =   3
      HeadingRowCount =   1
      HeadingColCount =   1
      TextAlignment   =   0
      WordWrap        =   0   'False
      Ellipsis        =   1
      HeadingBackColor=   -2147483633
      HeadingForeColor=   255
      HeadingTextAlignment=   5
      HeadingWordWrap =   0   'False
      HeadingEllipsis =   1
      GridLines       =   1
      HeadingGridLines=   2
      GridLinesColor  =   -2147483633
      HeadingGridLinesColor=   -2147483632
      EvenOddStyle    =   0
      ColorEven       =   -2147483628
      ColorOdd        =   -2147483624
      UserResizeAnimate=   1
      UserResizing    =   3
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      UserDragging    =   2
      UserHiding      =   2
      CellPadding     =   15
      CellBkgStyle    =   9
      CellBackColor   =   12632256
      CellForeColor   =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   1
      FocusRectColor  =   0
      FocusRectLineWidth=   1
      TabKeyBehavior  =   0
      EnterKeyBehavior=   0
      NavigationWrapMode=   1
      SkipReadOnly    =   0   'False
      DefaultColWidth =   1200
      DefaultRowHeight=   255
      CellsBorderColor=   0
      CellsBorderVisible=   -1  'True
      RowNumbering    =   0   'False
      EqualRowHeight  =   0   'False
      EqualColWidth   =   0   'False
      HScrollHeight   =   0
      VScrollWidth    =   0
      Format          =   "General"
      Appearance      =   2
      FitLastColumn   =   -1  'True
      SelectionMode   =   0
      MultiSelect     =   1
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      AllowEdit       =   -1  'True
      ScrollBarTips   =   0
      CellTips        =   0
      CellTipsDelay   =   1000
      SpecialMode     =   0
      OutlineLines    =   1
      CacheAllRecords =   -1  'True
      ColumnClickSort =   0   'False
      PreviewPaneColumn=   ""
      PreviewPaneType =   0
      PreviewPanePosition=   2
      PreviewPaneSize =   2000
      GroupIndentation=   225
      InactiveSelection=   1
      AutoScroll      =   -1  'True
      AutoResize      =   1
      AutoResizeHeadings=   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   0
      Caption         =   "Hedhja e Notave"
      ScrollTipColumn =   ""
      MaxRows         =   4194304
      MaxColumns      =   8192
      NewRowPos       =   1
      CustomBkgDraw   =   0
      AutoGroup       =   -1  'True
      GroupByBoxVisible=   0   'False
      GroupByBoxText  =   "Drag a column header here to group by that column"
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      StylesCollection=   $"frmSharpGrid.frx":003C
      ColumnsCollection=   $"frmSharpGrid.frx":1DD7
      ValueItems      =   $"frmSharpGrid.frx":30EC
   End
   Begin VB.Label lblDate 
      Caption         =   "Date:"
      Height          =   195
      Left            =   5280
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmSharpGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDil_Click()
 Unload Me
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
  With SGGrid1
      .Width = Me.Width - 300
'      ResizeControls Me, 0
   End With
End Sub

Private Sub Form_Load()
  
  
  'initializeGrid
  
  
  
End Sub


Private Sub initializeGrid()
   ' Initialize grid properties
   With SGGrid1
      Set .BkgPicture = LoadPicture(App.Path & "\bkgnd.jpg")
      
      .UserDragging = sgAllowColDrag
      .UserHiding = sgNoRowColHide
      .UserResizing = sgAllowColResizing
      
      .BkgPictureAlignment = sgPicAlignTile
      .FitLastColumn = True
      .Appearance = sg3DLight
      .SpecialMode = sgModeListBox
      .GridLines = sgGridLineFlat
      .GridLinesColor = vbBlue
      .OutlineLines = sgNoOutlineLines
      .CellsBorderVisible = True
      .AutoResize = sgAutoResizeRowsAndColumns
      .AllowDelete = True
      .AllowEdit = True
      .AllowAddNew = True
      .HeadingColCount = 0
      .HeadingGridLinesColor = vbYellow
      .HeadingGridLines = sgGridLineFlat
      .DataMode = sgUnbound
      .CellTips = sgCellTipsNone
      .ScrollBarTips = sgScrollTipsNone
      .CacheAllRecords = True
      .ColumnClickSort = False
      
      With .Styles("Normal")
         .BkgStyle = sgCellBkgNone
         .Font.Name = "Verdana"
         .Font.Size = 10
         .Padding = 18
      End With
      
      With .Styles("Heading")
         .BackColor = RGB(206, 48, 0)
         .BkgStyle = sgCellBkgSolid
         .ForeColor = vbWhite
         .Font.Name = "Verdana"
         .Font.Bold = True
         .Padding = 40
         .TextAlignment = sgAlignCenterCenter
      End With
      .Rows.At(0).Height = 290
      
      With .Styles("GroupHeader")
         .Font.Size = 9
         .BackColor = RGB(255, 207, 0)
         .BkgStyle = sgCellBkgSolid
         .Padding = 18
         .BorderColor = RGB(255, 207, 0)
         .Borders = sgCellBorderBottom
         .BorderSize = 1
      End With
      
      With .Styles("GroupFooter")
         .Font.Size = 9
         .BackColor = RGB(255, 255, 224)
         .BkgStyle = sgCellBkgSolid
         .Padding = 18
         .BorderColor = RGB(255, 207, 0)
         .Borders = sgCellBorderBottom
         .BorderSize = 1
      End With
            
      With .Styles("Selection")
         .BackColor = RGB(240, 128, 0)
         .ForeColor = vbWhite
         .BkgStyle = sgCellBkgSolid
      End With
      
      With .Styles("InactiveSelection")
         .BackColor = RGB(192, 192, 240)
         .ForeColor = vbBlack
         .BkgStyle = sgCellBkgSolid
      End With
      
      .GroupByBoxVisible = False
      
      .Caption = "KOT"
      With .Styles("Caption")
           .BackColor = vbBlue
      End With
      
   End With
    
   
   Dim objColumn As SGColumn
   ' Initialize columns and groupings
   With SGGrid1
        '.Rows.RemoveAll
        .DataRowCount = 0
        
        .Columns.RemoveAll True
       
      
       .Columns.Add ("Data")
      .Columns("Data").Hidden = True
      
      
      With .Columns.Add("EmriNxenesit")
        .Caption = "Emri Nxenesit"
        .Width = 2500
      End With
      
      With .Columns.Add("Matematike")  'lende id
           .Caption = "Matematike"
           .Width = 1000 'lende emer
      End With
      
      With .Columns.Add("Letersi")
           .Caption = "Letersi"
      End With
    

'      Dim grp1 As SGGroup
'
'      Set grp1 = .Groups.Add("Data", sgSortAscending, , True, False)
'
'
'         grp1.HeaderTextSource = sgGrpHdrCaptionAndValue
'       .RefreshGroups
'       .CollapseAll
       .RedrawEnabled = True

   End With
   loadTestData
End Sub


Private Sub loadTestData()
    
    SGGrid1.Rows.Add sgFormatCharSeparatedValue, "Melsi,9 ,9 , ", ""
    
    
    SGGrid1.RefreshGroups
    SGGrid1.CollapseAll
     SGGrid1.RedrawEnabled = True
End Sub

Private Sub Form_Resize()
  With SGGrid1
      .Width = Me.Width - 300

   End With
   
    txtDate.Left = SGGrid1.Left + SGGrid1.Width - txtDate.Width
    lblDate.Left = txtDate.Left
End Sub
