VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGridHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
' * Author           :
' * Web Site         :
' * E-Mail           :
' * Date             : 11/22/2004
' * Time             : 20:53
' * Module Name      : clsGridHandler
' * Module Filename  : clsGridHandler.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************




Option Explicit

Public Enum GridStyle
    EditGrid
    ListGrid
    ListGridSimple
    LupaGrid
    LupaMultiSelect
End Enum


Public Sub applyStyleGrid(Grid As SGGrid, Optional Style As GridStyle = LupaGrid)
   'ketu vendosen property te pergjitheshme
   ApplyCommonStyle Grid

   'pastaj te vecantat
   Select Case Style
      Case EditGrid:
         ApplyStyleEdit Grid

      Case ListGrid
         ApplyStyleList Grid

      Case ListGridSimple
         ApplyStyleListSimple Grid

      Case LupaGrid
         ApplyStyleLupa Grid

      Case LupaMultiSelect
         ApplyStyleLupaMulti Grid

   End Select

End Sub


Private Sub ApplyStyleList(Grid As SGGrid)

   With Grid
      .AllowAddNew = False
      .AllowDelete = False
      .AllowEdit = False
      .ColumnClickSort = True
      .GroupByBoxVisible = False
      .MultiSelect = sgMultiSelectNone
      .SpecialMode = sgModeListBox
      .SelectionMode = sgSelectionByRow
      .AutoSearch = True
   End With

End Sub


Private Sub ApplyStyleListSimple(Grid As SGGrid)

   With Grid
      .AllowAddNew = False
      .AllowDelete = False
      .AllowEdit = False
      .ColumnClickSort = False
      .GroupByBoxVisible = False
      .MultiSelect = sgMultiSelectNone
      .SpecialMode = sgModeListBox
      .SelectionMode = sgSelectionByRow
      .AutoSearch = True
   End With

End Sub


Private Sub ApplyStyleLupa(Grid As SGGrid)

   With Grid
      .AllowAddNew = False
      .AllowDelete = False
      .AllowEdit = False
      .UserDragging = sgAllowColDrag
      .ColumnClickSort = True
      .GroupByBoxVisible = True
      .MultiSelect = sgMultiSelectNone
      .SpecialMode = sgModeListBox
      .SelectionMode = sgSelectionByRow
      .AutoSearch = True
   End With

End Sub


Private Sub ApplyStyleLupaMulti(Grid As SGGrid)

   With Grid
      .AllowAddNew = False
      .AllowDelete = False
      .AllowEdit = False
      .UserDragging = sgAllowColDrag
      .ColumnClickSort = True
      .GroupByBoxVisible = True
      .SelectionMode = sgSelectionByRow
      .MultiSelect = sgMultiSelectExtended
      .SpecialMode = sgModeListBox
      .AutoSearch = True

   End With

End Sub


Private Sub ApplyStyleEdit(Grid As SGGrid)

   With Grid
      .AllowAddNew = True
      .AllowDelete = True
      .AllowEdit = True
      .EnterKeyBehavior = sgEnterKeyColumns
      .TabKeyBehavior = sgTabColumns
      .NavigationWrapMode = sgNavigationWrapNext
      .UserDragging = sgNoRowColDrag
      .ColumnClickSort = False
      .GroupByBoxVisible = False
      .MultiSelect = sgMultiSelectSimple
   End With

End Sub


Private Sub ApplyCommonStyle(Grid As SGGrid)

   With Grid
      .GroupByBoxText = "Zvarrisni nje kolone ketu per ti grupuar te dhenat"
      .AutoResize = sgNoAutoResize
      .AutoScroll = True
      '.BackColor
      .CellTips = sgCellTipsFloat
      .CellTipsDelay = 100

      .CellForeColor = vbBlue
      '.CellsBorderColor
      '.CellsBorderVisible
      '.ColorEven
      '.ColorOdd
      '.ColumnClickSort = False

      'Probleme kur behet set kjo property ne:
      '.DataMode = sgUnbound

      .EqualRowHeight = True

      '.EvenOddStyle = sgEvenOddRows
      .GridLines = sgGridLineFlat
      .RowNumbering = True
      '.ScrollBars = sgSBAll
      .SelectionMode = sgSelectionByRow
      .SkipReadOnly = True
      .TabKeyBehavior = sgTabColumns
      .FitLastColumn = True
      .EnterKeyBehavior = sgEnterKeyRows
   End With

End Sub


Public Function AddCol(Grid As SGGrid, Emri As String, strCaption As String, GjeresiNePerqind As Double, Optional TipiControllit As sgCellEditorType = sgCellTextBox, Optional TipiTeDhenes As sgDataType = sgtString) As SGColumn
   Dim Col              As SGColumn

   'Totali i perqindjes duhet 96%

   Set Col = Grid.Columns.Add(Emri)
   With Col
      .Caption = strCaption
      .DataType = TipiTeDhenes
      .Control.Type = TipiControllit
      .Width = CLng(Grid.Width * (GjeresiNePerqind / 100))
      '.Style.TextAlignment = sgAlignCenterCenter
      .HeadingStyle.TextAlignment = sgAlignCenterCenter
   End With

   Set AddCol = Col
   Set Col = Nothing

End Function


Public Sub mergeGrid(objGrid As SGGrid)
   With objGrid
      .Columns.At(0).MergeCells = sgMergeFree
   End With
End Sub


Public Sub applyStyleGrid1(objGrid As SGGrid, strCaption As String, editMode As Boolean)
   Dim objColumn        As SGColumn

   With objGrid

      .AutoResize = sgAutoResizeRowsAndColumns

      Set .BkgPicture = LoadPicture(App.Path & "\Images\bkgnd.jpg")

      .UserDragging = sgAllowColDrag
      .UserHiding = sgNoRowColHide
      .UserResizing = sgAllowColResizing

      .BkgPictureAlignment = sgPicAlignTile
      .FitLastColumn = True
      .Appearance = sg3DLight
      .SpecialMode = sgModeListBox
      .GridLines = sgGridLineFlat
      .GridLinesColor = vb3DShadow
      .OutlineLines = sgNoOutlineLines
      .CellsBorderVisible = False
      .AutoResize = sgAutoResizeRowsAndColumns
      .AllowEdit = editMode
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

      .Caption = strCaption

      With .Styles("Caption")
         .BackColor = &H8000000C
         .ForeColor = vbBlue
      End With

      '.Rows.RemoveAll
      .DataRowCount = 0
      .Columns.RemoveAll False

      .Columns.Add ("Data")
      .Columns("Data").Hidden = True

      .Columns.Add ("EmriNxenesit")

      .Columns.Add ("Matematike1")
      .Columns.Add ("Matematike2")
      .Columns("Matematike2").Width = 2500

      Set objColumn = .Columns("EmriNxenesit")
      objColumn.Caption = "Emri Nxenesit"
      objColumn.Style.WordWrap = True

      Set objColumn = .Columns("Matematike1")
      objColumn.Caption = "Matematike"
      objColumn.Style.WordWrap = True

      Set objColumn = .Columns("Matematike2")
      objColumn.Caption = "Matematike"


      .RedrawEnabled = True

   End With
   loadTestData objGrid
End Sub


Private Sub loadTestData(objGrid As SGGrid)

   objGrid.Rows.Add sgFormatCharSeparatedValue, "Melsi, Melsi,10,4", ","
   objGrid.Rows.Add sgFormatCharSeparatedValue, "Vali,Vali,5,7", ","
   objGrid.Rows.Add sgFormatCharSeparatedValue, "xxx,xxxx,6,8", ","
   objGrid.RefreshGroups
   objGrid.CollapseAll
   objGrid.RedrawEnabled = True
   
End Sub
