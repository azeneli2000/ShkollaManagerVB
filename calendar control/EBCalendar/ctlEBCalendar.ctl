VERSION 5.00
Begin VB.UserControl ctlEBCalendar 
   BackColor       =   &H80000005&
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2175
   ScaleHeight     =   3015
   ScaleWidth      =   2175
   ToolboxBitmap   =   "ctlEBCalendar.ctx":0000
   Begin VB.CommandButton cmdToday 
      Caption         =   "&Today"
      Height          =   315
      Left            =   660
      TabIndex        =   46
      Top             =   2580
      Width           =   855
   End
   Begin VB.PictureBox picInputBack 
      BackColor       =   &H80000005&
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   2115
      TabIndex        =   43
      Top             =   0
      Width           =   2175
      Begin VB.TextBox txtDate 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   15
         TabIndex        =   45
         Top             =   15
         Width           =   1815
      End
      Begin VB.CommandButton cmdDropDown 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   8.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   44
         Top             =   0
         Width           =   195
      End
   End
   Begin VB.Line linSep 
      BorderColor     =   &H80000015&
      Index           =   5
      X1              =   2160
      X2              =   0
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line linSep 
      BorderColor     =   &H80000015&
      Index           =   4
      X1              =   2160
      X2              =   -60
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line linSep 
      BorderColor     =   &H80000015&
      Index           =   3
      X1              =   2160
      X2              =   2160
      Y1              =   300
      Y2              =   3000
   End
   Begin VB.Line linSep 
      BorderColor     =   &H80000015&
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   300
      Y2              =   3000
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   35
      Left            =   60
      TabIndex        =   55
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   36
      Left            =   360
      TabIndex        =   54
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   37
      Left            =   660
      TabIndex        =   53
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   38
      Left            =   960
      TabIndex        =   52
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   39
      Left            =   1260
      TabIndex        =   51
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   40
      Left            =   1560
      TabIndex        =   50
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   41
      Left            =   1860
      TabIndex        =   49
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblMonthNext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1980
      TabIndex        =   48
      Top             =   360
      Width           =   195
   End
   Begin VB.Label lblMonthPrev 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   60
      TabIndex        =   47
      Top             =   360
      Width           =   195
   End
   Begin VB.Line linSep 
      BorderColor     =   &H8000000F&
      Index           =   2
      X1              =   60
      X2              =   2100
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   34
      Left            =   1860
      TabIndex        =   42
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   33
      Left            =   1560
      TabIndex        =   41
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   32
      Left            =   1260
      TabIndex        =   40
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   31
      Left            =   960
      TabIndex        =   39
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   30
      Left            =   660
      TabIndex        =   38
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   29
      Left            =   360
      TabIndex        =   37
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   28
      Left            =   60
      TabIndex        =   36
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   27
      Left            =   1860
      TabIndex        =   35
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   26
      Left            =   1560
      TabIndex        =   34
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   25
      Left            =   1260
      TabIndex        =   33
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   24
      Left            =   960
      TabIndex        =   32
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   23
      Left            =   660
      TabIndex        =   31
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   22
      Left            =   360
      TabIndex        =   30
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   21
      Left            =   60
      TabIndex        =   29
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   1860
      TabIndex        =   28
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   1560
      TabIndex        =   27
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   1260
      TabIndex        =   26
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   960
      TabIndex        =   25
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   660
      TabIndex        =   24
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   360
      TabIndex        =   23
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   60
      TabIndex        =   22
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   1860
      TabIndex        =   21
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   1560
      TabIndex        =   20
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   1260
      TabIndex        =   19
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   960
      TabIndex        =   18
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   660
      TabIndex        =   17
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   16
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   60
      TabIndex        =   15
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   1860
      TabIndex        =   14
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   13
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1260
      TabIndex        =   12
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   11
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   660
      TabIndex        =   10
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   9
      Top             =   960
      Width           =   255
   End
   Begin VB.Line linSep 
      BorderColor     =   &H8000000F&
      Index           =   1
      X1              =   60
      X2              =   2100
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   8
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblDOW 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sa"
      Height          =   255
      Index           =   6
      Left            =   1860
      TabIndex        =   7
      Top             =   660
      Width           =   255
   End
   Begin VB.Label lblDOW 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fr"
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   6
      Top             =   660
      Width           =   255
   End
   Begin VB.Label lblDOW 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Th"
      Height          =   255
      Index           =   4
      Left            =   1260
      TabIndex        =   5
      Top             =   660
      Width           =   255
   End
   Begin VB.Label lblDOW 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "We"
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   4
      Top             =   660
      Width           =   255
   End
   Begin VB.Label lblDOW 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tu"
      Height          =   255
      Index           =   2
      Left            =   660
      TabIndex        =   3
      Top             =   660
      Width           =   255
   End
   Begin VB.Label lblDOW 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mo"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   660
      Width           =   255
   End
   Begin VB.Label lblDOW 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Su"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   255
   End
   Begin VB.Label lblMonthYear 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      BackStyle       =   0  'Transparent
      Caption         =   "Month 0000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2235
   End
   Begin VB.Shape shpMonthYearBack 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   315
      Left            =   -60
      Top             =   300
      Width           =   2295
   End
End
Attribute VB_Name = "ctlEBCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[Description]
'   A simple date field with drop down calendar

'[Notes]
'   Ensure the form has enough room below the control to display the calendar
'   otherwise it will be clipped
'   Also ensure that the controls zorder is set to 0 (by right-clicking on the
'   control and selecting Bring To Front) otherwise the popup window appears
'   behind other controls.

'[Author]
'   Richard Allsebrook  <RA>    richardallsebrook@earlybirdmarketing.com

'[Contributions]
'   Wolfgang Wicker     <WW>    wolfgang.wicker@ggaweb.ch

'[History]
'   Version 1.1.0
'   Updates for internationalisation
'   Thanks to <WW> who pointed out the lack of internationalisation
'   support.
'   Updated the DrawCalendar routine to use WeekDay to find first day of month
'   and Initialise Function to use local day names for calendar day headings.

'   Version 1.0.0
'   Initial release

'[Declarations]
Option Explicit

Private flgCalVisible       As Boolean          'Is calcendar currently visible
Private dtCalDate           As Date             'Calendar page being displayed

'Property storage
Private dtDate              As Date             'Current date value
Private flgAutoSelect       As Boolean          'Do we auto select date field
                                                'on GotFocus

Private Sub cmdDropDown_Click()

'[Description]
'   Show/Hide Calendar

'[Code]

    'Toggle calendar visibility flag
    flgCalVisible = Not flgCalVisible
    
    If flgCalVisible Then
        'Redraw calendar to match current date setting
    
        If dtDate = 0 Then
            'No date currently entered - default to today
            dtCalDate = Now
            
        Else
            'Use current date
            dtCalDate = dtDate
        End If
        
        DrawCalendar
        
    End If
    
    'Display/Hide calendar
    UserControl_Resize
      
End Sub

Private Sub cmdToday_Click()

'[Description]
'   Set date to today and hide the calendar

'[Code]

    Me.Text = Now
    flgCalVisible = False
    UserControl_Resize
    
End Sub

Private Sub lblDay_Click(Index As Integer)

'[Description]
'   Set the date to that selected by the user and hide the calendar

'[Code]

    Me.Text = CDate(lblDay(Index).Caption & " " & lblMonthYear)
    flgCalVisible = False
    UserControl_Resize
    
End Sub

Private Sub lblMonthNext_Click()

'[Description]
'   Display the next month

'[Code]

    dtCalDate = DateSerial(Year(dtCalDate), Month(dtCalDate) + 1, Day(dtCalDate))
    DrawCalendar
    
End Sub

Private Sub lblMonthPrev_Click()

'[Description]
'   Display last previous month

'[Code]

    dtCalDate = DateSerial(Year(dtCalDate), Month(dtCalDate) - 1, Day(dtCalDate))
    DrawCalendar
    
End Sub

Private Sub txtDate_GotFocus()

'[Description]
'   If AutoSelect is enabled, select the fields contents

'[Code]

    If flgAutoSelect Then
    
        With txtDate
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        
    End If
    
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)

'[Description]
'   If a valid date has been entered, store it, otherwise cancel

'[Code]

    If IsDate(txtDate.Text) Then
        Me.Text = txtDate.Text
        
    Else
        Cancel = True
    End If
    
End Sub

Private Sub UserControl_ExitFocus()

'[Description]
'   If the control looses focus and the calendar is currently visible, hide
'   the calendar

'[Code]

    If flgCalVisible Then
        flgCalVisible = False
        UserControl_Resize
    End If
    
End Sub

Private Sub UserControl_Initialize()

'[Description]
'   Prepare the control

'[Declaration]
Dim intIndex As Integer
'[Code]

    flgCalVisible = False
    dtCalDate = Now
    
    For intIndex = 0 To 6
        lblDOW(intIndex).Caption = Left(Format(intIndex + 1, "ddd", vbSunday), 2)
    Next
        
End Sub

Private Sub UserControl_InitProperties()

'[Description]
'   Initialise the control's properties

'[Code]

    dtDate = Now
    flgAutoSelect = True
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'[Description]
'   Read the control's properties

'[Code]

    With PropBag
        Me.Text = .ReadProperty("Text", Now)
        flgAutoSelect = .ReadProperty("AutoSelect", True)
    End With
    
End Sub

Private Sub UserControl_Resize()

'[Description]
'   Redraw the control to match it's new dimentions

'[Code]

    With UserControl
    
        .Width = 2175 'fix width
        
        If flgCalVisible Then
            .Height = 3015 'drop down calendar
        Else
            .Height = 315 'hide calendar
        End If
        
    End With
    
End Sub

Private Function DrawCalendar()

'[Description]
'   Update the calendar to match the new dtCalDate

'[Declarations]
Dim intIndex                As Integer
Dim strDay                  As String
Dim intDay                  As Integer

    'Update header
    lblMonthYear = Format(dtCalDate, "MMMM YYYY")

    'Reset the day labels
    For intIndex = 0 To 41
        
        With lblDay(intIndex)
            .BackColor = vbWindowBackground
            .ForeColor = vbWindowText
            .Visible = False
            .FontBold = False
        End With
        
    Next
    
    'Find first day of month
    intIndex = Format("1 " & Month(dtCalDate) & " " & Year(dtCalDate), "w") - 1
    
    intDay = 1
    
    For intIndex = intIndex To intIndex + Day(DateSerial(Year(dtCalDate), Month(dtCalDate) + 1, 0)) - 1
        
        With lblDay(intIndex)
            .Caption = intDay
            
            If Format(intDay & " " & lblMonthYear, "dd mmm yyyy") = Format(Now, "dd mmm yyyy") Then
                .FontBold = True
            End If
            
            If Format(intDay & " " & lblMonthYear, "dd mmm yyyy") = Format(dtDate, "dd mmm yyyy") Then
                .BackColor = vbHighlight
                .ForeColor = vbHighlightText
            End If
            
            .Visible = True
        End With
        
        intDay = intDay + 1
    Next
    
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'[Description]
'   Store updated properties

'[Code]

    With PropBag
        .WriteProperty "Text", dtDate, Now
        .WriteProperty "AutoSelect", flgAutoSelect, True
    End With
    
End Sub

Public Property Get Text() As Date

'[Description]
'   Return the current date

'[Code]

    If dtDate = 0 Then
        'no current date
        Text = Now
        
    Else
        Text = Format(dtDate, "dd mmm yyyy")
    End If
    
End Property

Public Property Let Text(NewValue As Date)

'[Description]
'   Set the current date

'[Code]

    dtDate = NewValue
    dtCalDate = NewValue
    
    If dtDate = 0 Then
        'No date
        txtDate.Text = ""
        
    Else
        txtDate.Text = Format(dtDate, "dd mmm yyyy")
    End If
    
    If flgCalVisible Then
        'Update the calendar
        DrawCalendar
    End If
    
    PropertyChanged "Text"
    
End Property

Public Property Get AutoSelect() As Boolean

'[Description]
'   Return the current AutoSelect property

'[Code]

    AutoSelect = flgAutoSelect
    
End Property

Public Property Let AutoSelect(NewValue As Boolean)

'[Description]
'   Set the AutoSelect property

'[Code]

    flgAutoSelect = NewValue
    PropertyChanged "AutoSelect"
    
End Property
