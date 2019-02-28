VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl ucFlexGrid 
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4860
   ScaleHeight     =   3285
   ScaleWidth      =   4860
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   3300
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3135
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   5530
      _Version        =   393216
   End
End
Attribute VB_Name = "ucFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private UsingMouse As Boolean  ' Flag for using the Mouse in the Grid.
'Event Declarations:
Event Click() 'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,Click
Attribute Click.VB_Description = "Fired when the user presses and releases the mouse button over the control."
Event DblClick() 'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,DblClick
Attribute DblClick.VB_Description = "Fired when the user double-clicks the mouse over the control."
Event EnterCell() 'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,EnterCell
Attribute EnterCell.VB_Description = "Fired before the cursor enters a cell."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Text2,Text2,-1,KeyDown
Attribute KeyDown.VB_Description = "Fired when the user pushes a key."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Text2,Text2,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Text2,Text2,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event LeaveCell() 'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,LeaveCell
Attribute LeaveCell.VB_Description = "Fired after the cursor leaves a cell."
Event RowColChange() 'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,RowColChange
Attribute RowColChange.VB_Description = "Fired when the current cell changes."
Event SelChange() 'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,SelChange
Attribute SelChange.VB_Description = "Fired when the selected range of cells changes."

Private Type SelProperties
    Locked As Boolean
End Type

Private mVarColumn() As SelProperties

Public Property Get ColLocked(index As Integer) As Boolean
    ColLocked = mVarColumn(index).Locked
End Property

Public Property Let ColLocked(index As Integer, bLocked As Boolean)
    mVarColumn(index).Locked = bLocked
End Property

Private Sub UserControl_Initialize()
    Text2.ZOrder 0
    MSFlexGrid1.TabStop = False
    
End Sub

Private Sub Text2_GotFocus()
   MSFlexGrid1.Text = Text2.Text
   If MSFlexGrid1.Col >= MSFlexGrid1.Cols Then MSFlexGrid1.Col = 1
   ChangeCellText
End Sub

Private Sub MSFlexGrid1_EnterCell()  ' Assign cell value to the textbox
    RaiseEvent EnterCell
    If Not mVarColumn(MSFlexGrid1.Col).Locked Then
       Text2.BackColor = &HC0FFFF
       Text2.Locked = False
    Else
       Text2.BackColor = vbRed
       Text2.Locked = True
    End If
       
    Text2.Text = MSFlexGrid1.Text
    
End Sub

Private Sub MSFlexGrid1_LeaveCell()
    RaiseEvent LeaveCell
' Assign textbox value to grid

   MSFlexGrid1.Text = Text2.Text
   Text2.Text = ""
   If MSFlexGrid1.Row = MSFlexGrid1.Rows - 1 And MSFlexGrid1.Text <> "" Then
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
   End If
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, _
                                  x As Single, y As Single)
   UsingMouse = True
   MSFlexGrid1.Text = Text2.Text
   ChangeCellText
End Sub

Private Sub Text2_LostFocus()
   If UsingMouse = True Then
      UsingMouse = False
      Exit Sub
   End If
   
   If MSFlexGrid1.Col <= MSFlexGrid1.Cols - 2 Then
      MSFlexGrid1.Col = MSFlexGrid1.Col + 1
      ChangeCellText
   Else
      If MSFlexGrid1.Row + 1 < MSFlexGrid1.Rows Then
        MSFlexGrid1.Row = MSFlexGrid1.Row + 1
        MSFlexGrid1.Col = 1
        ChangeCellText
      Else
        MSFlexGrid1.Row = 1
        MSFlexGrid1.Col = 1
        ChangeCellText
      End If
   End If
End Sub

Public Sub ChangeCellText() ' Move Textbox to active cell.
   Dim iRow As Integer
   Dim iCol As Integer
   Text2.Move MSFlexGrid1.Left + MSFlexGrid1.CellLeft, _
   MSFlexGrid1.Top + MSFlexGrid1.CellTop, _
   MSFlexGrid1.CellWidth, MSFlexGrid1.CellHeight
   Text2.SetFocus
   Text2.ZOrder 0
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    
    MSFlexGrid1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,CellHeight
Public Property Get CellHeight() As Long
Attribute CellHeight.VB_Description = "Returns the height of the current cell, in Twips."
    CellHeight = MSFlexGrid1.CellHeight
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,CellLeft
Public Property Get CellLeft() As Long
Attribute CellLeft.VB_Description = "Returns the left position of the current cell, in twips"
    CellLeft = MSFlexGrid1.CellLeft
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,CellPicture
Public Property Get CellPicture() As Picture
Attribute CellPicture.VB_Description = "Returns/sets an image to be displayed in the current cell or in a range of cells."
    Set CellPicture = MSFlexGrid1.CellPicture
End Property

Public Property Set CellPicture(ByVal New_CellPicture As Picture)
    Set MSFlexGrid1.CellPicture = New_CellPicture
    PropertyChanged "CellPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,CellTop
Public Property Get CellTop() As Long
Attribute CellTop.VB_Description = "Returns or sets the top position of the current cell, in twips"
    CellTop = MSFlexGrid1.CellTop
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,CellWidth
Public Property Get CellWidth() As Long
Attribute CellWidth.VB_Description = "Returns the width of the current cell, in twips"
    CellWidth = MSFlexGrid1.CellWidth
End Property

Private Sub MSFlexGrid1_Click()
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,ColAlignment
Public Property Get ColAlignment(ByVal index As Long) As Integer
Attribute ColAlignment.VB_Description = "Returns/sets the alignment of data in a column. Not available at design time (except indirectly through the FormatString property)."
    ColAlignment = MSFlexGrid1.ColAlignment(index)
End Property

Public Property Let ColAlignment(ByVal index As Long, ByVal New_ColAlignment As Integer)
    MSFlexGrid1.ColAlignment(index) = New_ColAlignment
    PropertyChanged "ColAlignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,ColData
Public Property Get ColData(ByVal index As Long) As Long
Attribute ColData.VB_Description = "Array of long integer values with one item for each row (RowData) and for each column (ColData) of the FlexGrid. Not available at design time."
    ColData = MSFlexGrid1.ColData(index)
End Property

Public Property Let ColData(ByVal index As Long, ByVal New_ColData As Long)
    MSFlexGrid1.ColData(index) = New_ColData
    PropertyChanged "ColData"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,ColIsVisible
Public Property Get ColIsVisible(ByVal index As Long) As Boolean
Attribute ColIsVisible.VB_Description = "Returns True if the specified column is visible."
    ColIsVisible = MSFlexGrid1.ColIsVisible(index)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,ColPos
Public Property Get ColPos(ByVal index As Long) As Long
Attribute ColPos.VB_Description = "Returns the distance in Twips between the upper-left corner of the control and the upper-left corner of a specified column."
    ColPos = MSFlexGrid1.ColPos(index)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,ColPosition
'Public Property Get ColPosition(ByVal index As Long) As Long
'    ColPosition = MSFlexGrid1.ColPosition(index)
'End Property
'
'Public Property Let ColPosition(ByVal index As Long, ByVal New_ColPosition As Long)
'    MSFlexGrid1.ColPosition(index) = New_ColPosition
'    PropertyChanged "ColPosition"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,Cols
Public Property Get Cols() As Long
Attribute Cols.VB_Description = "Determines the total number of columns or rows in a FlexGrid."
    Cols = MSFlexGrid1.Cols
End Property

Public Property Let Cols(ByVal New_Cols As Long)
    MSFlexGrid1.Cols() = New_Cols
    ReDim mVarColumn(New_Cols) As SelProperties
    PropertyChanged "Cols"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,ColWidth
Public Property Get ColWidth(ByVal index As Long) As Long
Attribute ColWidth.VB_Description = "Determines the width of the specified column in Twips. Not available at design time."
    ColWidth = MSFlexGrid1.ColWidth(index)
End Property

Public Property Let ColWidth(ByVal index As Long, ByVal New_ColWidth As Long)
    MSFlexGrid1.ColWidth(index) = New_ColWidth
    PropertyChanged "ColWidth"
End Property

Private Sub MSFlexGrid1_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = MSFlexGrid1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    MSFlexGrid1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,FixedCols
Public Property Get FixedCols() As Long
Attribute FixedCols.VB_Description = "Returns/sets the total number of fixed (non-scrollable) columns or rows for a FlexGrid."
    FixedCols = MSFlexGrid1.FixedCols
End Property

Public Property Let FixedCols(ByVal New_FixedCols As Long)
    MSFlexGrid1.FixedCols() = New_FixedCols
    PropertyChanged "FixedCols"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,FixedRows
Public Property Get FixedRows() As Long
Attribute FixedRows.VB_Description = "Returns/sets the total number of fixed (non-scrollable) columns or rows for a FlexGrid."
    FixedRows = MSFlexGrid1.FixedRows
End Property

Public Property Let FixedRows(ByVal New_FixedRows As Long)
    MSFlexGrid1.FixedRows() = New_FixedRows
    PropertyChanged "FixedRows"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,FixedAlignment
Public Property Get FixedAlignment(ByVal index As Long) As Integer
Attribute FixedAlignment.VB_Description = "Returns/sets the alignment of data in the fixed cells of a column."
    FixedAlignment = MSFlexGrid1.FixedAlignment(index)
End Property

Public Property Let FixedAlignment(ByVal index As Long, ByVal New_FixedAlignment As Integer)
    MSFlexGrid1.FixedAlignment(index) = New_FixedAlignment
    PropertyChanged "FixedAlignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,GridLines
Public Property Get GridLines() As GridLineSettings
Attribute GridLines.VB_Description = "Returns/sets the type of lines that should be drawn between cells."
    GridLines = MSFlexGrid1.GridLines
End Property

Public Property Let GridLines(ByVal New_GridLines As GridLineSettings)
    MSFlexGrid1.GridLines() = New_GridLines
    PropertyChanged "GridLines"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,GridLinesFixed
Public Property Get GridLinesFixed() As GridLineSettings
Attribute GridLinesFixed.VB_Description = "Returns/sets the type of lines that should be drawn between cells."
    GridLinesFixed = MSFlexGrid1.GridLinesFixed
End Property

Public Property Let GridLinesFixed(ByVal New_GridLinesFixed As GridLineSettings)
    MSFlexGrid1.GridLinesFixed() = New_GridLinesFixed
    PropertyChanged "GridLinesFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,GridLineWidth
Public Property Get GridLineWidth() As Integer
Attribute GridLineWidth.VB_Description = "Returns/sets the width in Pixels of the gridlines for the control."
    GridLineWidth = MSFlexGrid1.GridLineWidth
End Property

Public Property Let GridLineWidth(ByVal New_GridLineWidth As Integer)
    MSFlexGrid1.GridLineWidth() = New_GridLineWidth
    PropertyChanged "GridLineWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,GridColorFixed
Public Property Get GridColorFixed() As Long
Attribute GridColorFixed.VB_Description = "Returns/sets the color used to draw the lines between FlexGrid cells."
    GridColorFixed = MSFlexGrid1.GridColorFixed
End Property

Public Property Let GridColorFixed(ByVal New_GridColorFixed As Long)
    MSFlexGrid1.GridColorFixed() = New_GridColorFixed
    PropertyChanged "GridColorFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,GridColor
Public Property Get GridColor() As Long
Attribute GridColor.VB_Description = "Returns/sets the color used to draw the lines between FlexGrid cells."
    GridColor = MSFlexGrid1.GridColor
End Property

Public Property Let GridColor(ByVal New_GridColor As Long)
    MSFlexGrid1.GridColor() = New_GridColor
    PropertyChanged "GridColor"
End Property

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,MergeCells
Public Property Get MergeCells() As MergeCellsSettings
Attribute MergeCells.VB_Description = "Returns/sets whether cells with the same contents should be grouped in a single cell spanning multiple rows or columns."
    MergeCells = MSFlexGrid1.MergeCells
End Property

Public Property Let MergeCells(ByVal New_MergeCells As MergeCellsSettings)
    MSFlexGrid1.MergeCells() = New_MergeCells
    PropertyChanged "MergeCells"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,MergeCol
Public Property Get MergeCol(ByVal index As Long) As Boolean
Attribute MergeCol.VB_Description = "Returns/sets which rows (columns) should have their contents merged when the MergeCells property is set to a value other than 0 - Never."
    MergeCol = MSFlexGrid1.MergeCol(index)
End Property

Public Property Let MergeCol(ByVal index As Long, ByVal New_MergeCol As Boolean)
    MSFlexGrid1.MergeCol(index) = New_MergeCol
    PropertyChanged "MergeCol"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,MergeRow
Public Property Get MergeRow(ByVal index As Long) As Boolean
Attribute MergeRow.VB_Description = "Returns/sets which rows (columns) should have their contents merged when the MergeCells property is set to a value other than 0 - Never."
    MergeRow = MSFlexGrid1.MergeRow(index)
End Property

Public Property Let MergeRow(ByVal index As Long, ByVal New_MergeRow As Boolean)
    MSFlexGrid1.MergeRow(index) = New_MergeRow
    PropertyChanged "MergeRow"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,MouseCol
Public Property Get MouseCol() As Long
Attribute MouseCol.VB_Description = "Returns/sets over which row (column) the mouse pointer is. Not available at design time."
    MouseCol = MSFlexGrid1.MouseCol
End Property

Private Sub MSFlexGrid1_RowColChange()
    RaiseEvent RowColChange
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,RowData
Public Property Get RowData(ByVal index As Long) As Long
Attribute RowData.VB_Description = "Array of long integer values with one item for each row (RowData) and for each column (ColData) of the FlexGrid. Not available at design time."
    RowData = MSFlexGrid1.RowData(index)
End Property

Public Property Let RowData(ByVal index As Long, ByVal New_RowData As Long)
    MSFlexGrid1.RowData(index) = New_RowData
    PropertyChanged "RowData"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,RowHeight
Public Property Get RowHeight(ByVal index As Long) As Long
Attribute RowHeight.VB_Description = "Returns/sets the height of the specified row in Twips. Not available at design time."
    RowHeight = MSFlexGrid1.RowHeight(index)
End Property

Public Property Let RowHeight(ByVal index As Long, ByVal New_RowHeight As Long)
    MSFlexGrid1.RowHeight(index) = New_RowHeight
    PropertyChanged "RowHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,RowHeightMin
Public Property Get RowHeightMin() As Long
Attribute RowHeightMin.VB_Description = "Returns/sets a minimum row height for the entire control, in Twips."
    RowHeightMin = MSFlexGrid1.RowHeightMin
End Property

Public Property Let RowHeightMin(ByVal New_RowHeightMin As Long)
    MSFlexGrid1.RowHeightMin() = New_RowHeightMin
    PropertyChanged "RowHeightMin"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,RowIsVisible
Public Property Get RowIsVisible(ByVal index As Long) As Boolean
Attribute RowIsVisible.VB_Description = "Returns True if the specified row is visible."
    RowIsVisible = MSFlexGrid1.RowIsVisible(index)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,RowPos
Public Property Get RowPos(ByVal index As Long) As Long
Attribute RowPos.VB_Description = "Returns the distance in Twips between the upper-left corner of the control and the upper-left corner of a specified row."
    RowPos = MSFlexGrid1.RowPos(index)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,RowPosition
'Public Property Get RowPosition(ByVal index As Long) As Long
'    RowPosition = MSFlexGrid1.RowPosition(index)
'End Property
'
'Public Property Let RowPosition(ByVal index As Long, ByVal New_RowPosition As Long)
'    MSFlexGrid1.RowPosition(index) = New_RowPosition
'    PropertyChanged "RowPosition"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,Rows
Public Property Get Rows() As Long
Attribute Rows.VB_Description = "Determines the total number of columns or rows in a FlexGrid."
    Rows = MSFlexGrid1.Rows
End Property

Public Property Let Rows(ByVal New_Rows As Long)
    MSFlexGrid1.Rows() = New_Rows
    PropertyChanged "Rows"
End Property

Private Sub MSFlexGrid1_SelChange()
    RaiseEvent SelChange
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,SelectionMode
Public Property Get SelectionMode() As SelectionModeSettings
Attribute SelectionMode.VB_Description = "Returns/sets whether a FlexGrid should allow regular cell selection, selection by rows, or selection by columns."
    SelectionMode = MSFlexGrid1.SelectionMode
End Property

Public Property Let SelectionMode(ByVal New_SelectionMode As SelectionModeSettings)
    MSFlexGrid1.SelectionMode() = New_SelectionMode
    PropertyChanged "SelectionMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text2,Text2,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = Text2.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    Text2.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text2,Text2,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = Text2.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    Text2.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text2,Text2,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = Text2.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    Text2.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,TextMatrix
Public Property Get TextMatrix(ByVal Row As Long, ByVal Col As Long) As String
Attribute TextMatrix.VB_Description = "Returns/sets the text contents of an arbitrary cell (row/col subscripts)."
    TextMatrix = MSFlexGrid1.TextMatrix(Row, Col)
End Property

Public Property Let TextMatrix(ByVal Row As Long, ByVal Col As Long, ByVal New_TextMatrix As String)
    MSFlexGrid1.TextMatrix(Row, Col) = New_TextMatrix
    PropertyChanged "TextMatrix"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,TextMatrix
Public Property Get Text() As String
    Text = MSFlexGrid1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    MSFlexGrid1.Text = New_Text
    PropertyChanged "Text"
End Property


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim index As Integer

    Set CellPicture = PropBag.ReadProperty("CellPicture", Nothing)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    MSFlexGrid1.ColAlignment(index) = PropBag.ReadProperty("ColAlignment" & index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    MSFlexGrid1.ColData(index) = PropBag.ReadProperty("ColData" & index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    MSFlexGrid1.ColPosition(index) = PropBag.ReadProperty("ColPosition" & index, 0)
    MSFlexGrid1.Cols = PropBag.ReadProperty("Cols", 2)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    MSFlexGrid1.ColWidth(index) = PropBag.ReadProperty("ColWidth" & index, 0)
    MSFlexGrid1.Enabled = PropBag.ReadProperty("Enabled", True)
    MSFlexGrid1.FixedCols = PropBag.ReadProperty("FixedCols", 1)
    MSFlexGrid1.FixedRows = PropBag.ReadProperty("FixedRows", 1)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    MSFlexGrid1.FixedAlignment(index) = PropBag.ReadProperty("FixedAlignment" & index, 0)
    MSFlexGrid1.GridLines = PropBag.ReadProperty("GridLines", 1)
    MSFlexGrid1.GridLinesFixed = PropBag.ReadProperty("GridLinesFixed", 2)
    MSFlexGrid1.GridLineWidth = PropBag.ReadProperty("GridLineWidth", 1)
    MSFlexGrid1.GridColorFixed = PropBag.ReadProperty("GridColorFixed", 0)
    MSFlexGrid1.GridColor = PropBag.ReadProperty("GridColor", 12632256)
    MSFlexGrid1.MergeCells = PropBag.ReadProperty("MergeCells", 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    MSFlexGrid1.MergeCol(index) = PropBag.ReadProperty("MergeCol" & index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    MSFlexGrid1.MergeRow(index) = PropBag.ReadProperty("MergeRow" & index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    MSFlexGrid1.RowData(index) = PropBag.ReadProperty("RowData" & index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    MSFlexGrid1.RowHeight(index) = PropBag.ReadProperty("RowHeight" & index, 0)
    MSFlexGrid1.RowHeightMin = PropBag.ReadProperty("RowHeightMin", 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    MSFlexGrid1.RowPosition(index) = PropBag.ReadProperty("RowPosition" & index, 0)
    MSFlexGrid1.Rows = PropBag.ReadProperty("Rows", 2)
    MSFlexGrid1.SelectionMode = PropBag.ReadProperty("SelectionMode", 0)
    Text2.SelLength = PropBag.ReadProperty("SelLength", 0)
    Text2.SelStart = PropBag.ReadProperty("SelStart", 0)
    Text2.SelText = PropBag.ReadProperty("SelText", "")
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    'MSFlexGrid1.TextMatrix(Row, Col) = PropBag.ReadProperty("TextMatrix" & Index, "")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim index As Integer

    Call PropBag.WriteProperty("CellPicture", CellPicture, Nothing)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Call PropBag.WriteProperty("ColAlignment" & index, MSFlexGrid1.ColAlignment(index), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Call PropBag.WriteProperty("ColData" & index, MSFlexGrid1.ColData(index), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    'Call PropBag.WriteProperty("ColPosition" & Index, MSFlexGrid1.ColPosition(Index), 0)
    Call PropBag.WriteProperty("Cols", MSFlexGrid1.Cols, 2)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Call PropBag.WriteProperty("ColWidth" & index, MSFlexGrid1.ColWidth(index), 0)
    Call PropBag.WriteProperty("Enabled", MSFlexGrid1.Enabled, True)
    Call PropBag.WriteProperty("FixedCols", MSFlexGrid1.FixedCols, 1)
    Call PropBag.WriteProperty("FixedRows", MSFlexGrid1.FixedRows, 1)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Call PropBag.WriteProperty("FixedAlignment" & index, MSFlexGrid1.FixedAlignment(index), 0)
    Call PropBag.WriteProperty("GridLines", MSFlexGrid1.GridLines, 1)
    Call PropBag.WriteProperty("GridLinesFixed", MSFlexGrid1.GridLinesFixed, 2)
    Call PropBag.WriteProperty("GridLineWidth", MSFlexGrid1.GridLineWidth, 1)
    Call PropBag.WriteProperty("GridColorFixed", MSFlexGrid1.GridColorFixed, 0)
    Call PropBag.WriteProperty("GridColor", MSFlexGrid1.GridColor, 12632256)
    Call PropBag.WriteProperty("MergeCells", MSFlexGrid1.MergeCells, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Call PropBag.WriteProperty("MergeCol" & index, MSFlexGrid1.MergeCol(index), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Call PropBag.WriteProperty("MergeRow" & index, MSFlexGrid1.MergeRow(index), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Call PropBag.WriteProperty("RowData" & index, MSFlexGrid1.RowData(index), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Call PropBag.WriteProperty("RowHeight" & index, MSFlexGrid1.RowHeight(index), 0)
    Call PropBag.WriteProperty("RowHeightMin", MSFlexGrid1.RowHeightMin, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    'Call PropBag.WriteProperty("RowPosition" & Index, MSFlexGrid1.RowPosition(Index), 0)
    Call PropBag.WriteProperty("Rows", MSFlexGrid1.Rows, 2)
    Call PropBag.WriteProperty("SelectionMode", MSFlexGrid1.SelectionMode, 0)
    Call PropBag.WriteProperty("SelLength", Text2.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", Text2.SelStart, 0)
    Call PropBag.WriteProperty("SelText", Text2.SelText, "")
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    'Call PropBag.WriteProperty("TextMatrix" & Index, MSFlexGrid1.TextMatrix(Row, Col), "")
End Sub

