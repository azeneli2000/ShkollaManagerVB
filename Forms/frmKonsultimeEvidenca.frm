VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{04759352-AB08-11D5-863E-0010B563AF37}#1.0#0"; "Cgridv11.ocx"
Begin VB.Form frmKonsultimeEvidenca 
   Caption         =   "Konsultime - Evidenca"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   165.894
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   268.817
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboIndeksi 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      ForeColor       =   &H80000004&
      Height          =   315
      Left            =   5760
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   8775
      Begin VB.Frame fraOptions 
         Caption         =   "Nota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1335
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   3255
         Begin VB.OptionButton optSemestri1 
            Caption         =   "Semestrale I"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optSemestri2 
            Caption         =   "Semestrale II"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   840
            Width           =   1695
         End
         Begin VB.OptionButton optVjetore 
            Caption         =   "Vjetore"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   23
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.ComboBox cboKlasa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000004&
         Height          =   315
         Left            =   4800
         TabIndex        =   21
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cboVitiShkollor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000004&
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdKerko 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "frmKonsultimeEvidenca.frx":0000
         DownPicture     =   "frmKonsultimeEvidenca.frx":6442
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6840
         Picture         =   "frmKonsultimeEvidenca.frx":C884
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Klasa 
         Caption         =   "Klasa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   28
         Top             =   690
         Width           =   615
      End
      Begin VB.Label lblVitiShkollor 
         Caption         =   "Viti Shkollor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Indeksi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   26
         Top             =   720
         Width           =   735
      End
   End
   Begin MSComCtl2.FlatScrollBar fsbVertikali 
      Height          =   6015
      Left            =   14910
      TabIndex        =   10
      Top             =   2040
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   10610
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1245184
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid amzaEmerMbiemer 
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   2040
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   -2147483626
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.FlatScrollBar fsbHorizontali 
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   7800
      Visible         =   0   'False
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Arrows          =   65536
      Orientation     =   1245185
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmKonsultimeEvidenca.frx":12CC6
      DownPicture     =   "frmKonsultimeEvidenca.frx":16200
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      Picture         =   "frmKonsultimeEvidenca.frx":1973A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9000
      Width           =   2415
   End
   Begin VB.Frame fraInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   12480
      TabIndex        =   4
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton cmdWebsiste 
         BackColor       =   &H80000009&
         Caption         =   "Website"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdEmail 
         BackColor       =   &H80000009&
         Caption         =   "E-mail"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmKonsultimeEvidenca.frx":1CC74
         Stretch         =   -1  'True
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblWebsite 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblShkolla 
         Appearance      =   0  'Flat
         Caption         =   "Shkolla jo publike Nr 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmKonsultimeEvidenca.frx":23BE9
      DownPicture     =   "frmKonsultimeEvidenca.frx":27123
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      Picture         =   "frmKonsultimeEvidenca.frx":2A65D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9000
      Width           =   2175
   End
   Begin VB.CommandButton cmdDalje 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmKonsultimeEvidenca.frx":2DB97
      DownPicture     =   "frmKonsultimeEvidenca.frx":33FD9
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      Picture         =   "frmKonsultimeEvidenca.frx":3A41B
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9000
      Width           =   2535
   End
   Begin Cgridv11.Cgrid provimet 
      Height          =   1095
      Left            =   8520
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1931
      GridMode        =   0
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   16
      FixedColumnVisible=   -1  'True
      FixedRowVisible =   -1  'True
      ScrollBarV      =   0   'False
      ScrollBarh      =   0   'False
      Appearance3D    =   0   'False
      CellEditColor   =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid emrat 
      Height          =   4815
      Left            =   0
      TabIndex        =   12
      Top             =   2295
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      TabIndex        =   16
      Top             =   1920
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Lendet 
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   2040
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   -2147483626
      FixedCols       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid notat 
      Height          =   4815
      Left            =   2880
      TabIndex        =   13
      Top             =   2280
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   8493
      _Version        =   393216
      ScrollBars      =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      TabIndex        =   17
      Top             =   120
      Width           =   15135
   End
End
Attribute VB_Name = "frmKonsultimeEvidenca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objektGabimi As New clsErrorHandler
Dim objPrintim As Object
Dim objPrintForm As New frmPrintEvidenca
Dim objPrintClass As Object
Dim TabelNotash(20, 3) As String
Dim Gabim As Boolean


Private Sub cboIndeksi_Click()
    Pastro
End Sub

Private Sub cboKlasa_Click()
    Pastro
End Sub



Private Sub cboVitiShkollor_Click()
    Pastro
End Sub

Private Sub cmdDalje_Click()
   Unload Me
   Set active_form = Nothing
End Sub

Private Sub cmdDil_Click()
    CallHelp indeksHelp
End Sub


Private Sub cmdEmail_Click()
    If Not OpenEmailProgram(email) Then
        MsgBox "Ju nuk keni asnje program te instaluar per te derguar email", vbExclamation, "Gabim ne dergimin e emailit"
    End If
End Sub

Private Sub cmdKerko_Click()
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      objectInitialization VISUALIZO_EVIDENCE
      If notat.Visible = True Then
        If optVjetore.Value Then
            If notat.Cols > 15 Then
                fsbHorizontali.Visible = True
            'FlatScrollBar1.Top = 5000
            Else
                fsbHorizontali.Visible = False
            End If
        Else
            If notat.Cols > 40 Then
                fsbHorizontali.Visible = True
                'FlatScrollBar1.Top = 5000
            Else
                fsbHorizontali.Visible = False
            End If
        End If
        If notat.Rows > 20 Then
            fsbVertikali.Visible = True
        Else
            fsbVertikali.Visible = False
        End If
      End If
        cmdKerko.Enabled = True
   End If
End Sub

Private Sub cmdOK_Click()
   'gridaNotat.Visible = False
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      objectInitialization VISUALIZO_EVIDENCE
      If scrbar = True Then
         FlatScrollBar1.Visible = True
         'FlatScrollBar1.Top = 5000
      Else
        FlatScrollBar1.Visible = False
      End If
      If scrbarver = True Then
        FlatScrollBar2.Visible = True
      Else
        FlatScrollBar2.Visible = False
      End If
      cmdOK.Enabled = True
   End If
End Sub


Private Sub cmdKopjo_Click()
    Clipboard.Clear
    Clipboard.SetText notat.Clip
End Sub

Private Sub cmdPrint_Click()
   Dim llojPrintim      As Integer
   Dim nResult As Integer
   objPrintForm.show vbModal
   llojPrintim = objPrintForm.nota
   Unload objPrintForm
   Select Case llojPrintim
      Case 0:
         Exit Sub
      Case 1:
         Set objPrintim = CreateObject("PrintimComponent.clsPrintim")
        If objPrintim.PrinterIsInstalled = False Then
            Exit Sub
        End If
         'objPrintim.FormatPrinter "vbPRPSA3"
         'If objPrintim.Gabim = True Then
         '   nResult = MsgBox("Printeri juaj nuk ju lejon qe printimi te kryhet ne format A3. Per nje lende" _
         '   & " ju mund te printoni deri ne 10 nota", vbOKCancel + vbInformation, "Informacion per printerin")
         '   If nResult = vbOK Then
         '       NoteGrida
         '   Else
         '       Exit Sub
         '   End If
         'Else
            NoteGrida 'A3
         'End If
      Case 2:
         objectInitialization PRINTO_EVIDENCA
      Case 4:
         PrintoFormatSem
      Case 3:
         PrintoFormatNx
   End Select
End Sub



Private Sub cmdWebsiste_Click()
    If website <> "" Then
        GoToWeb website
    Else
        MsgBox "Ju nuk e keni dhene adresen e faqes tuaj te web-it.", vbInformation
    End If
End Sub


Private Sub emrat_Click()
Dim nr As Integer
nr = emrat.row
notat.row = nr
notat.RowSel = nr

notat.SetFocus
End Sub

Private Sub Form_Load()
   loadForm Me
   mbushKomboBox
   viti
   Image1.Picture = LoadPicture(adresaLogo)
   lblShkolla.Caption = emerShkolla
   'txtEmri(1).Visible = False
   fsbVertikali.Visible = False
   fsbHorizontali.Visible = False
   cmdPrint.Enabled = False
   notat.Visible = False
   emrat.Visible = False
   Lendet.Visible = False
   amzaEmerMbiemer.Visible = False
End Sub


Private Sub Form_Resize()
 'With SGGrid1
     ' .Width = Me.Width - 400
      
      '.Height = Me.Height
   'End With
   'Frame1.Width = Me.Width - 400
   'fraInfo.Left = Me.Left + Me.Width - fraInfo.Width - 200
End Sub


Private Sub Form_Activate()
emrat.SelectionMode = flexSelectionByRow
notat.SelectionMode = flexSelectionByRow
 ' With SGGrid1
   '   .Width = Me.Width - 400
      '.Height = Me.Height
  ' End With
End Sub



Private Sub fsbHorizontali_Change()
   
   Dim t                As Integer
   Dim vlera            As Integer
   If Not optVjetore.Value Then
      
      t = 53
      vlera = fsbHorizontali.Value
      Lendet.left = 53 - vlera
      Lendet.Width = 40 * CDbl(300 / 56.7) + vlera
      notat.left = 53 - vlera
      notat.Width = 40 * CDbl(300 / 56.7) + vlera
   Else
      
      t = 53
      vlera = fsbHorizontali.Value
      Lendet.left = 53 - vlera
      Lendet.Width = 15 * CDbl(700 / 56.7) + vlera
      notat.left = 53 - vlera
      notat.Width = 15 * CDbl(700 / 56.7) + vlera
   End If

End Sub

Private Sub fsbHorizontali_Scroll()
   Dim t                As Integer
   Dim vlera            As Integer
   If Not optVjetore.Value Then
      
      t = 53
      vlera = fsbHorizontali.Value
      Lendet.left = 53 - vlera
      Lendet.Width = 40 * CDbl(300 / 56.7) + vlera
      notat.left = 53 - vlera
      notat.Width = 40 * CDbl(300 / 56.7) + vlera
   Else
      
      t = 53
      vlera = fsbHorizontali.Value
      Lendet.left = 53 - vlera
      Lendet.Width = 15 * CDbl(700 / 56.7) + vlera
      notat.left = 53 - vlera
      notat.Width = 15 * CDbl(700 / 56.7) + vlera
   End If
End Sub

Private Sub fsbVertikali_Change()
    Dim t As Integer
    t = 35 + CDbl(300 / 56.7)
    
    notat.top = t - fsbVertikali.Value
    emrat.top = t - fsbVertikali.Value
    notat.Height = 20 * CDbl(300 / 56.7) + fsbVertikali.Value
    emrat.Height = 20 * CDbl(300 / 56.7) + fsbVertikali.Value
End Sub

Private Sub fsbVertikali_Scroll()
     Dim t As Integer
    t = 35 + CDbl(300 / 56.7)
    
    notat.top = t - fsbVertikali.Value
    emrat.top = t - fsbVertikali.Value
    notat.Height = 20 * CDbl(300 / 56.7) + fsbVertikali.Value
    emrat.Height = 20 * CDbl(300 / 56.7) + fsbVertikali.Value
End Sub
Public Sub inicializoGridaNotat(numer_rreshtash As Integer, numer_shtyllash As Integer)
    
    gridaNotat.Width = 14575
    gridaNotat.Rows = numer_rreshtash
    gridaNotat.Cols = numer_shtyllash
    gridaNotat.FixedRows = 1
    gridaNotat.FixedCols = 0
    gridaNotat.row = 0
    gridaNotat.col = 0
    gridaNotat.ColWidth(0) = 2000
    gridaNotat.Text = "Numri i amzes"
    gridaNotat.col = 1
    gridaNotat.ColWidth(1) = 2000
    gridaNotat.Text = "Emri i nxenesit"
    gridaNotat.col = 2
    gridaNotat.ColWidth(2) = 2000
    gridaNotat.Text = "Matematike"
    gridaNotat.col = 3
    gridaNotat.ColWidth(3) = 2000
    gridaNotat.Text = "Fizike"
    gridaNotat.col = 4
    gridaNotat.ColWidth(4) = 7500
    gridaNotat.Text = "Letersi"
   
End Sub

Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
End Sub

Private Sub mbushKomboBox()
    cboVitiShkollor.AddItem "1990-1991"
    cboVitiShkollor.AddItem "1991-1992"
    cboVitiShkollor.AddItem "1992-1993"
    cboVitiShkollor.AddItem "1993-1994"
    cboVitiShkollor.AddItem "1994-1995"
    cboVitiShkollor.AddItem "1995-1996"
    cboVitiShkollor.AddItem "1996-1997"
    cboVitiShkollor.AddItem "1997-1998"
    cboVitiShkollor.AddItem "1998-1999"
    cboVitiShkollor.AddItem "1999-2000"
    cboVitiShkollor.AddItem "2000-2001"
    cboVitiShkollor.AddItem "2001-2002"
    cboVitiShkollor.AddItem "2002-2003"
    cboVitiShkollor.AddItem "2003-2004"
    cboVitiShkollor.AddItem "2004-2005"
    cboVitiShkollor.AddItem "2005-2006"
    cboVitiShkollor.AddItem "2006-2007"
    cboVitiShkollor.AddItem "2007-2008"
    cboVitiShkollor.AddItem "2008-2009"
    cboVitiShkollor.AddItem "2009-2010"
    cboVitiShkollor.AddItem "2010-2011"
    cboVitiShkollor.AddItem "2011-2012"
    cboVitiShkollor.AddItem "2012-2012"
    cboVitiShkollor.AddItem "2013-2014"
    cboVitiShkollor.AddItem "2014-2015"
    cboVitiShkollor.AddItem "2015-2016"
    cboVitiShkollor.AddItem "2016-2017"
    cboVitiShkollor.AddItem "2017-2018"
    cboVitiShkollor.AddItem "2018-2019"
    cboVitiShkollor.AddItem "2019-2020"
    
    cboIndeksi.AddItem "A"
    cboIndeksi.AddItem "B"
    cboIndeksi.AddItem "C"
    cboIndeksi.AddItem "D"
    cboIndeksi.AddItem "E"
    cboIndeksi.AddItem "F"
    
    cboKlasa.AddItem "1"
    cboKlasa.AddItem "2"
    cboKlasa.AddItem "3"
    cboKlasa.AddItem "4"
    cboKlasa.AddItem "5"
    cboKlasa.AddItem "6"
    cboKlasa.AddItem "7"
    cboKlasa.AddItem "8"
    cboKlasa.AddItem "9"
    cboKlasa.AddItem "10"
    cboKlasa.AddItem "11"
    cboKlasa.AddItem "12"
End Sub

Private Sub viti()
   Dim d, muai, viti    As String
   d = Now
   Dim M, v             As Integer
   muai = DateTime.Month(DateTime.Now)
   viti = DateTime.Year(DateTime.date)
   M = Val(muai)
   v = Val(viti)
   cboVitiShkollor.Text = gjej_vitin(M, v)
End Sub


Private Sub NoteGrida()
    Dim nrShtylla As Integer, nrRreshta As Integer, nrKolonash As Integer
    Dim nrShtyllaPergj As Integer
    Dim i, j, k, p, nResult, l As Integer
    Dim totalPages As Integer
    Dim currPage As Integer
    Dim tekst As String
    Dim Portret As Boolean, fontbold As Boolean
    Dim show As Boolean
    Dim gjeresiKolone As Long
    nrShtylla = notat.Cols / Lendet.Cols
    nrRreshta = emrat.Rows
    
    If optVjetore.Value Then
        gjeresiKolone = 35
    Else
        If nrShtylla = 7 Then
            gjeresiKolone = 35
        Else
            gjeresiKolone = nrShtylla * 5
        End If
    End If
    nrKolonash = Fix(230 / gjeresiKolone)
    
    If notat.Rows > 32 Then
        M = 32
        ReDim tabelFaqeDy(notat.Rows - 32, nrKolonash)
        objPrintim.PrintEvidenca tabelFaqeDy
        Exit Sub
    End If

    Gabim = False
    
    Portret = False
    objPrintim.PrintFont "Times New Roman"
    objPrintim.OrientimFaqe Portret
    If objPrintim.Gabim = True Then
        MsgBox "Printeri juaj nuk lejon qe te nderrohet ne menyre automatike formati ne Landscape. " & _
            "Nderrojeni formatin nga Portrait ne Landscape te preferencat e printerit tuaj.", vbCritical + vbOKOnly, _
            "Gabim ne Printim."
        Exit Sub
    End If
    
    'objPrintim.PrintLeft 6, 22, "Emri Mbiemri", True, 10
    ' Nese perdoruesi anullon printimin
    If objPrintim.Gabim = True Then
        Exit Sub
    End If
    PrintoFormatPergj
    If Gabim = True Then
        Exit Sub
    End If
    'objPrintim.PrintTabele 5, 21, 45, 5, nrRreshta + 1, 1
    currPage = 1
    If Lendet.Cols Mod nrKolonash = 0 Then
        totalPages = Lendet.Cols / nrKolonash
    Else
        totalPages = Fix(Lendet.Cols / nrKolonash) + 1
    End If
    p = 0
    For i = 0 To Lendet.Cols - 1
        If p = 0 Then
            For j = 0 To emrat.Rows - 1
                objPrintim.PrintTabele 5, 21, 45, 5, nrRreshta + 1, 1
                If objPrintim.Gabim = True Then
                    Exit Sub
                End If
                objPrintim.PrintLeft 10, 22, "Emri  Mbiemri", True, 10, False
                emrat.col = 2
                emrat.row = j
                objPrintim.PrintLeft 6, 27 + (5 * j), emrat.Text, False, 10, False
            Next
            
        End If
        objPrintim.PrintTabele 50 + p * gjeresiKolone, 21, gjeresiKolone, 5, nrRreshta + 1, 1
        Lendet.row = 0
        Lendet.col = i
        objPrintim.PrintLeft 51 + p * gjeresiKolone, 22, Lendet.Text, True, 10
        For j = 0 To notat.Rows - 1
            If j > 35 Then
                Exit Sub
            End If
            l = i * nrShtylla
            k = p * nrShtylla
            Do While k < (p + 1) * nrShtylla
                notat.col = l
                notat.row = j
                If notat.CellForeColor = &H800000 Or notat.CellForeColor = &HC0& Then
                    fontbold = True
                Else
                    fontbold = False
                End If
                If optVjetore.Value = True Then
                    objPrintim.PrintLeft 52 + k * 11.667, 27 + 5 * j, notat.Text, fontbold, 10
                Else
                    objPrintim.PrintLeft 50.2 + k * 5, 27 + 5 * j, notat.Text, fontbold, 10
                End If
                k = k + 1
                l = l + 1
            Loop
        Next
        p = p + 1
        If p >= nrKolonash Then
            objPrintim.PrintLeft 285, 197, CStr(currPage) & "/" & CStr(totalPages), True, 10
            If currPage = totalPages Then
                Exit For
            End If
            objPrintim.NewPage
            currPage = currPage + 1
            'If i + 1 = Lendet.Cols Then
            '    Exit For
            'End If
            objPrintim.OrientimFaqe False
            p = 0
            PrintoFormatPergj
            If Gabim = True Then
                Exit Sub
            End If
        End If
    Next
    objPrintim.PrintLeft 285, 197, CStr(currPage) & "/" & totalPages, True, 10
    objPrintim.EndDoc
End Sub

' Metode qe ben formatimin dhe printimin e nje qelize te grides se notave
Private Sub FormatCgrid1(Qeliza As String, PosX As Long, PosY As Long, formati As String)
    Dim i, j, k As Integer
    Dim notaDhjete As Integer
    Dim nrLendet As Integer
    i = 1
    j = 0
    notaDhjete = 0
    If formati = "A3" Then
        nrLendet = 14
    Else
        nrLendet = 9
    End If
    Do While i < Len(Qeliza)
        ' Nese ne gride jane me shume se 10 nota atehere dil nga metoda
        If j > nrLendet Then
            Exit Sub
        End If
        'Ne varesi te vlerave qe ndodhen ne qelizen e grides behet edhe formatimi i printimit
        If Mid(Qeliza, i, 2) = "||" Then
            If Mid(Qeliza, i + 2, 2) = "10" Then
                k = 2
            Else
                k = 1
            End If
            'TabelNotash(1, 1) = CInt(Mid(Qeliza, i + 2, k))
            objPrintim.PrintLeft PosX + (3 * j), PosY, Mid(Qeliza, i + 2, k), True, 6, False
            i = i + k + 4
            j = j + 1
            notaDhjete = k - 1
        ElseIf Mid(Qeliza, i, 1) = "|" Then
            If Mid(Qeliza, i + 1, 2) = "10" Then
                k = 2
            Else
                k = 1
            End If
            objPrintim.PrintLeft PosX + (3 * j), PosY, Mid(Qeliza, i + 1, k), False, 6, True
            i = i + 3
            j = j + 1
            notaDhjete = k - 1
        ElseIf Mid(Qeliza, i, 1) = " " Then
            i = i + 1
        Else
            If Mid(Qeliza, i, 2) = "10" Then
                k = 2
            Else
                k = 1
            End If
            objPrintim.PrintLeft PosX + (3 * j), PosY, Mid(Qeliza, i, k), False, 6, False
            i = i + k
            j = j + 1
            notaDhjete = k - 1
        End If
    Loop
End Sub

Private Sub notat_Click()
Dim nr As Integer
nr = notat.row
emrat.row = nr
emrat.RowSel = nr
emrat.SetFocus
End Sub

Private Sub optSemestri1_Click()
    Pastro
End Sub

Private Sub Pastro()
    notat.Visible = False
    emrat.Visible = False
    amzaEmerMbiemer.Visible = False
    Lendet.Visible = False
    fsbVertikali.Visible = False
    fsbHorizontali.Visible = False
    scrbar = False
    scrbarver = False
    cmdPrint.Enabled = False
End Sub

Private Sub optSemestri2_Click()
    Pastro
End Sub

Private Sub optVjetore_Click()
    Pastro
    
End Sub

Private Sub formatoGrida(grida As MSHFlexGrid, lartesia As Long)

   Dim l                As Long
   Dim i                As Integer
   i = grida.Rows
   l = i * CLng(grida.RowHeight(0)) + 50
   If l < lartesia Then
      grida.Height = l
   Else
      grida.Height = lartesia
   End If

End Sub
Private Sub formatoGridaGjeresi(grida As MSHFlexGrid, gjeresia As Long)

   Dim l                As Long
   Dim i                As Integer
   i = grida.Cols
   l = i * CLng(grida.ColWidth(0)) + 50
   If l < gjeresia Then
      grida.Width = l
   Else
      grida.Width = gjeresia
   End If

End Sub

Private Sub ngjyrosGriden(grida As MSHFlexGrid)

   Dim i                As Integer
   Dim j                As Integer
   For i = 0 To grida.Rows - 1
      If i Mod 2 = 1 Then
         For j = 0 To grida.Cols - 1
            grida.col = j
            grida.row = i
            grida.CellBackColor = &HE0E0E0
         Next j
      End If
   Next i
End Sub

Private Sub NoteGridaA3()
    Dim gridHeight, GridWidth As Integer
    Dim i, j, k, p, nResult, maksimum As Integer
    Dim tekst As String
    Dim Portret As Boolean
    Dim mbrapa As Boolean
    ' Numri i rreshtave dhe i shtyllave per gridat
    GridWidth = CInt(Lendet.Width / 56)
    GridWidth = notat.Cols / Lendet.Cols
    gridHeight = emrat.Rows
    mbrapa = True
    Portret = False
    objPrintim.PrintFont "Times New Roman"
    objPrintim.OrientimFaqe Portret
    If objPrintim.Gabim = True Then
        MsgBox "Printeri juaj nuk lejon qe te nderrohet ne menyre automatike formati ne Landscape. " & _
            "Nderrojeni formatin nga Portrait ne Landscape te preferencat e printerit tuaj.", vbCritical + vbOKOnly, _
            "Gabim ne Printim."
        Exit Sub
    End If
    
    objPrintim.PrintLeft 5, 10, "Emri Mbiemri", True, 8
    ' Nese perdoruesi anullon printimin
    If objPrintim.Gabim = True Then
        Exit Sub
    End If
    'objPrintim.PrintLine 42, 10, 42, (gridHeight * 6) + 14
    ' Printo emrat e nxenesve te lendeve dhe notat perkatese
    p = 1
    For i = 0 To GridWidth - 1
        Lendet.row = 0
        Lendet.col = i
        objPrintim.PrintLeft 5 + (42 * p), 10, Lendet.Text, False, 8
        k = 1
        j = 0
        Do While j < gridHeight
            'If show = True Then
                emrat.row = j
                emrat.col = 2
                objPrintim.PrintLeft 5, 10 + (6 * k), emrat.Text, True, 8
            'End If
            notat.row = j
            notat.col = i
            FormatCgrid1 notat.Text, 5 + (42 * p), 10 + (6 * k), "A3"
            objPrintim.PrintLine 5, 8 + (6 * k), 5 + (42 * (p + 1)), 8 + (6 * k)
            'objPrintim.PrintLine 5, 14 + (6 * k), 15 + (27.8 * (p + 1)), 14 + (6 * k)
            objPrintim.PrintLine 4 + (42 * p), 10, 4 + (42 * p), (gridHeight * 6) + 14
            objPrintim.PrintLine 4 + (42 * (p + 1)), 10, 4 + (42 * (p + 1)), (gridHeight * 6) + 14
            k = k + 1
            'If (20 + (7 * j)) > 200 Then
            '    objPrintim.EndDoc
            '    nResult = MsgBox("Ju lutemi futeni fleten nga ana tjeter per te printuar" & _
            '        " dhe lendet e tjera.", vbInformation + vbOKCancel, "Printimi i evidencave")
            '    If nResult = vbCancel Then
            '        Exit Sub
            '    End If
            '    objPrintim.NewPage
            '    k = 1
            'End If
            j = j + 1
            'Printer.Print
        Loop
        p = p + 1
        If p > 9 And mbrapa Then
            objPrintim.EndDoc
            nResult = MsgBox("Ju lutemi futeni fleten nga ana tjeter per te printuar" & _
                " edhe lendet e tjera.", vbInformation + vbOKCancel, "Printimi i evidencave")
            If nResult = vbCancel Then
                Exit Sub
            End If
            'objPrintim.NewPage
            p = 1
            mbrapa = False
        ElseIf p > 10 Then
            objPrintim.EndDoc
            nrResult = MsgBox("Fusni nje flete tjeter per te vazhduar printimin e notave", _
            vbOKCancel + vbInformation, "Printimi i evidencave")
            If nResult = vbCancel Then
                Exit Sub
            End If
            p = 0
        End If
    Next
    objPrintim.EndDoc
End Sub

Private Sub PrintoFormatPergj()
    objPrintClass.PrintLeft 15, 7, emerShkolla, True, 12, False, True
    If objPrintClass.Gabim = True Then
        Gabim = True
        Exit Sub
    End If
    objPrintClass.PrintLeft 70, 7, "Klasa ", True, 12, False
    objPrintClass.PrintLeft 82, 7, Space(2) & cboKlasa.Text & " " & cboIndeksi.Text & Space(2), True, 12, False, True
    objPrintClass.PrintLeft 200, 7, "Viti shkollor", True, 12, False
    objPrintClass.PrintLeft 225, 7, Space(2) & cboVitiShkollor.Text & Space(2), True, 12, False, True
    If qytetiShkolla <> "" Then
        objPrintClass.PrintLeft 15, 190, qytetiShkolla & " me " & FormatDateTime(Now, vbShortDate), True, 12, False, True
    End If
    If Image1 <> 0 Then
        objPrintClass.PrintPicture 135, 3, Image1, 16, 33
    End If
    objPrintClass.PrintLeft 200, 190, "Mesuesi kujdestar  ______________________", True, 12, False
End Sub


Private Sub PrintoFormatSem()
    Set objPrintClass = CreateObject("PrintimComponent.clsPrintim")
    If objPrintClass.PrinterIsInstalled = False Then
        Exit Sub
    End If
    Dim italik As Boolean
    Dim i, j As Integer
    Dim semesterString As String
    Dim tekst As String
    objPrintClass.PrintFont "Times New Roman"
    objPrintClass.OrientimFaqe True
    If objPrintClass.Gabim = True Then
        MsgBox "Printeri juaj nuk lejon qe te nderrohet ne menyre automatike formati ne Landscape. " & _
            "Nderrojeni formatin nga Portrait ne Landscape te preferencat e printerit tuaj.", vbCritical + vbOKOnly, _
            "Gabim ne Printim."
        Exit Sub
    End If
    objPrintClass.FormatPrinter "vbPPRSA4"
    If optSemestri1 Then
        semesterString = "parë"
    ElseIf optSemestri2 Then
        semesterString = "dytë"
    ElseIf Me.optVjetore Then
        FormatNotaPerf
        Exit Sub
    End If
    For i = 0 To emrat.Rows - 1
        emrat.row = i
        emrat.col = 2
        tekst = "Nxënësi " & emrat.Text & "ne klasen '" & cboKlasa.Text & " " & cboIndeksi.Text _
        & "' për semestrin e " & semesterString & " është vlerësuar me këto nota:"
        objPrintClass.PrintLeft 20, 88, tekst, False, 12
        If objPrintClass.Gabim = True Then
            Exit Sub
        End If
        PrintoFormatPergj1
        objPrintClass.PrintTabele 30, 105, 11, 6, 21, 1
        objPrintClass.PrintTabele 41, 105, 57, 6, 21, 1
        objPrintClass.PrintTabele 98, 105, 100, 6, 21, 1
        For j = 0 To Lendet.Cols - 1
            Lendet.row = 0
            Lendet.col = j
            objPrintClass.PrintLeft 42, 106 + (6 * j), Lendet.Text, False, 12
            b = 1
            For k = j * (notat.Cols / Lendet.Cols) To (j + 1) * (notat.Cols / Lendet.Cols) - 1
                notat.row = i
                notat.col = k
                objPrintClass.PrintLeft 93 + (6 * b), 106.5 + (6 * j), notat.Text, False, 12, False
                b = b + 1
            Next
        Next
        objPrintClass.NewPage
    Next
    objPrintClass.EndDoc
    Exit Sub
End Sub

Private Sub PrintoFormatNx()
    Set objPrintClass = CreateObject("PrintimComponent.clsPrintim")
    If objPrintClass.PrinterIsInstalled = False Then
        Exit Sub
    End If
    Dim italik As Boolean
    Dim i, j As Integer
    Dim semesterString As String
    Dim tekst As String
    objPrintClass.PrintFont "Times New Roman"
    objPrintClass.OrientimFaqe True
    If objPrintClass.Gabim = True Then
        MsgBox "Printeri juaj nuk lejon qe te nderrohet ne menyre automatike formati ne Landscape. " & _
            "Nderrojeni formatin nga Portrait ne Landscape te preferencat e printerit tuaj.", vbCritical + vbOKOnly, _
            "Gabim ne Printim."
        Exit Sub
    End If
    objPrintClass.FormatPrinter "vbPPRSA4"
    If optSemestri1 Then
        semesterString = "parë"
    ElseIf optSemestri2 Then
        semesterString = "dytë"
    ElseIf optVjetore Then
        FormatNotaPerfNje
        Exit Sub
    End If
    emrat.row = emrat.RowSel
    i = emrat.row
    emrat.col = 2
        tekst = "Nxënësi " & emrat.Text & " ne klasen '" & cboKlasa.Text & " " & cboIndeksi.Text _
        & "' për semestrin e " & semesterString & " është vlerësuar me këto nota:"
        objPrintClass.PrintLeft 20, 88, tekst, False, 12
        If objPrintClass.Gabim = True Then
            Exit Sub
        End If
        PrintoFormatPergj1
        objPrintClass.PrintTabele 30, 105, 11, 6, 21, 1
        objPrintClass.PrintTabele 41, 105, 57, 6, 21, 1
        objPrintClass.PrintTabele 98, 105, 100, 6, 21, 1
        For j = 0 To Lendet.Cols - 1
            Lendet.row = 0
            Lendet.col = j
            objPrintClass.PrintLeft 42, 106 + (6 * j), Lendet.Text, False, 12
            b = 1
            For k = j * (notat.Cols / Lendet.Cols) To (j + 1) * (notat.Cols / Lendet.Cols) - 1
                notat.row = i
                notat.col = k
                objPrintClass.PrintLeft 93 + (6 * b), 106.5 + (6 * j), notat.Text, False, 12, False
                b = b + 1
            Next
        Next
        objPrintClass.EndDoc
End Sub


Private Sub PrintoFormatPergj1()
    objPrintClass.PrintLeft 5, 17, lblShkolla.Caption, True, 12
    objPrintClass.PrintLeft 130, 17, "Adresa:", True, 12
    objPrintClass.PrintLeft 130, 23, adresaShkolla, True, 12
    objPrintClass.PrintLeft 20, 23, qytetiShkolla, True, 12
    objPrintClass.PrintLeft 130, 29, "Tel: " & telefoniShkolla, True, 12
    objPrintClass.PrintLeft 130, 35, "website: " & website, True, 12
    objPrintClass.PrintLeft 130, 41, "Email: " & email, True, 12
    If qytetiShkolla <> "" Then
        objPrintClass.PrintLeft 11, 261, qytetiShkolla & " me " & FormatDateTime(Now, vbShortDate), True, 12
    End If
    If Me.optVjetore.Value = True Then
        If PrintNota = 1 Then
            objPrintClass.PrintLeft 116, 261, "Drejtori  ____________________", True, 12
        Else
            objPrintClass.PrintLeft 116, 261, "Mesuesi kujdestar  ___________________", True, 12
        End If
    Else
        objPrintClass.PrintLeft 116, 261, "Mesuesi kujdestar  ___________________", True, 12
    End If
    If Image1 <> 0 Then
        objPrintClass.PrintPicture 81, 3, Image1, 40, 40
    End If
End Sub

Private Sub FormatNotaPerf()
    Set objPrintClass = CreateObject("PrintimComponent.clsPrintim")
    If objPrintClass.PrinterIsInstalled = False Then
        Exit Sub
    End If
    Dim tekst As String
    Dim i, j As Integer
    objPrintClass.PrintFont "Times New Roman"
    objPrintClass.OrientimFaqe True
    If objPrintClass.Gabim = True Then
        MsgBox "Printeri juaj nuk lejon qe te nderrohet ne menyre automatike formati ne Portrait. " & _
            "Nderrojeni formatin nga Landscape ne Portrait te preferencat e printerit tuaj.", vbCritical + vbOKOnly, _
            "Gabim ne Printim."
        Exit Sub
    End If
    Dim nx As Integer
    For nx = 0 To Me.emrat.Rows - 1
        emrat.row = nx
        emrat.col = 2
        objPrintClass.PrintLeft 90, 63, "VERTETIM", True, 12, False, True
        If objPrintClass.Gabim = True Then
            Exit Sub
        End If
        PrintoFormatPergj2
        Dim emri As String
        emri = emrat.Text
        tekst = "Vërtetojmë se nxënësi " & emrat.Text & _
            " me numër amze " & emrat.TextMatrix(nx, 1) & " në vitin shkollor " & cboVitiShkollor.Text
        objPrintClass.PrintLeft 30, 73, tekst, False, 12, False, False
        objPrintClass.PrintLeft 30, 79, " ka kryer klasën e " & MerrKlase & " dhe ka keto perfundime", False, 12
        ' Vendosja e tabelave qe do printohen
        objPrintClass.PrintTabele 45, 105, 11, 14, 1, 1
        objPrintClass.PrintTabele 56, 105, 57, 14, 1, 1
        objPrintClass.PrintTabele 113, 105, 57, 7, 1, 1
        objPrintClass.PrintTabele 113, 112, 19, 7, 1, 3
        objPrintClass.PrintTabele 45, 119, 11, 6, 20, 1
        objPrintClass.PrintTabele 56, 119, 57, 6, 20, 1
        objPrintClass.PrintTabele 113, 119, 19, 6, 20, 3
    
        objPrintClass.PrintLeft 46, 112, "Nr.", True, 12
        objPrintClass.PrintLeft 68, 112, "Lëndët", True, 12
        objPrintClass.PrintLeft 133, 107, "Notat", True, 12
        objPrintClass.PrintLeft 115, 114, "Semestri I", True, 10
        objPrintClass.PrintLeft 134, 114, "Semestri II", True, 10
        objPrintClass.PrintLeft 153, 114, "Vjetore", True, 10
        For i = 1 To 20
            objPrintClass.PrintLeft 46, 114.5 + (6 * i), CStr(i), False, 12, False
        Next
        i = 1
        k = 0
        For i = 0 To Lendet.Cols - 1
            objPrintClass.PrintLeft 57, 120.5 + (6 * i), Lendet.TextMatrix(0, i), False, 12, False
            objPrintClass.PrintLeft 117, 120.5 + (6 * i), notat.TextMatrix(nx, k), False, 12, False
            objPrintClass.PrintLeft 136, 120.5 + (6 * i), notat.TextMatrix(nx, k + 1), False, 12, False
            objPrintClass.PrintLeft 156, 120.5 + (6 * i), notat.TextMatrix(nx, k + 2), False, 12, False
            k = k + 3
        Next
        objPrintClass.NewPage
    Next
    objPrintClass.EndDoc
End Sub

Private Sub FormatNotaPerfNje()
    Set objPrintClass = CreateObject("PrintimComponent.clsPrintim")
    If objPrintClass.PrinterIsInstalled = False Then
        Exit Sub
    End If
    Dim tekst As String
    Dim i, j As Integer
    objPrintClass.PrintFont "Times New Roman"
    objPrintClass.OrientimFaqe True
    If objPrintClass.Gabim = True Then
        MsgBox "Printeri juaj nuk lejon qe te nderrohet ne menyre automatike formati ne Portrait. " & _
            "Nderrojeni formatin nga Landscape ne Portrait te preferencat e printerit tuaj.", vbCritical + vbOKOnly, _
            "Gabim ne Printim."
        Exit Sub
    End If
    Dim nx As Integer
    nx = emrat.RowSel
    emrat.row = nx
    emrat.col = 2
    objPrintClass.PrintLeft 90, 63, "VERTETIM", True, 12, False, True
    If objPrintClass.Gabim = True Then
        Exit Sub
    End If
    PrintoFormatPergj2
    tekst = "Vërtetojmë se nxënësi " & emrat.Text & _
        " me numër amze " & emrat.TextMatrix(nx, 1) & " në vitin shkollor " & cboVitiShkollor.Text
    objPrintClass.PrintLeft 30, 73, tekst, False, 12, False, False
    objPrintClass.PrintLeft 30, 79, " ka kryer klasën e " & MerrKlase & " dhe ka keto perfundime", False, 12
    ' Vendosja e tabelave qe do printohen
        objPrintClass.PrintTabele 45, 105, 11, 14, 1, 1
        objPrintClass.PrintTabele 56, 105, 57, 14, 1, 1
        objPrintClass.PrintTabele 113, 105, 57, 7, 1, 1
        objPrintClass.PrintTabele 113, 112, 19, 7, 1, 3
        objPrintClass.PrintTabele 45, 119, 11, 6, 20, 1
        objPrintClass.PrintTabele 56, 119, 57, 6, 20, 1
        objPrintClass.PrintTabele 113, 119, 19, 6, 20, 3
    
        objPrintClass.PrintLeft 46, 112, "Nr.", True, 12
        objPrintClass.PrintLeft 68, 112, "Lëndët", True, 12
        objPrintClass.PrintLeft 133, 107, "Notat", True, 12
        objPrintClass.PrintLeft 115, 114, "Semestri I", True, 10
        objPrintClass.PrintLeft 134, 114, "Semestri II", True, 10
        objPrintClass.PrintLeft 153, 114, "Vjetore", True, 10
    For i = 1 To 20
        objPrintClass.PrintLeft 51, 114.5 + (6 * i), CStr(i), False, 12, False
    Next
    i = 1
    k = 0
    For i = 0 To Lendet.Cols - 1
        objPrintClass.PrintLeft 57, 120.5 + (6 * i), Lendet.TextMatrix(0, i), False, 12, False
        objPrintClass.PrintLeft 117, 120.5 + (6 * i), notat.TextMatrix(nx, k), False, 12, False
        objPrintClass.PrintLeft 136, 120.5 + (6 * i), notat.TextMatrix(nx, k + 1), False, 12, False
        objPrintClass.PrintLeft 156, 120.5 + (6 * i), notat.TextMatrix(nx, k + 2), False, 12, False
        k = k + 3
    Next
    objPrintClass.EndDoc
End Sub
' Funksion qe merr si parameter nje numer dhe kthen ate numer ne
' shkronja qe i perkojne numrit ordinal te formatuar sipas defteses
Private Function MerrKlase() As String
    Dim numer As Integer
    On Error GoTo Gabimi
        numer = CInt(Me.cboKlasa.Text)
        If numer < 9 Then
            MerrKlase = OrdinalInLetters(numer)
        Else
            numer = numer - 8
            If numer = 0 Then
                MerrKlase = ""
            End If
            MerrKlase = OrdinalInLetters(numer)
        End If
        Exit Function
Gabimi:
        If Err.Number <> 0 Then
            MerrKlase = ""
        End If
End Function

' Merr si parameter nje numer dhe kthen ekuivalentin e tij ne shkronja
Private Function OrdinalInLetters(numer As Integer) As String
    Dim letter As String
    Select Case numer
    Case 1
        OrdinalInLetters = "parë"
    Case 2
        OrdinalInLetters = "dytë"
    Case 3
        OrdinalInLetters = "tretë"
    Case 4
        OrdinalInLetters = "katërt"
    Case 5
        OrdinalInLetters = "pestë"
    Case 6
        OrdinalInLetters = "gjashtë"
    Case 7
        OrdinalInLetters = "shtatë"
    Case 8
        OrdinalInLetters = "tetë"
    Case 9
        OrdinalInLetters = "nëntë"
    End Select
End Function

Private Sub PrintoFormatPergj2()
    objPrintClass.PrintLeft 5, 17, lblShkolla.Caption, True, 12
    objPrintClass.PrintLeft 130, 17, "Adresa:", True, 12
    objPrintClass.PrintLeft 130, 23, adresaShkolla, True, 12
    objPrintClass.PrintLeft 20, 23, qytetiShkolla, True, 12
    objPrintClass.PrintLeft 130, 29, "Tel: " & telefoniShkolla, True, 12
    objPrintClass.PrintLeft 130, 35, "website: " & website, True, 12
    objPrintClass.PrintLeft 130, 41, "Email: " & email, True, 12
    If qytetiShkolla <> "" Then
        objPrintClass.PrintLeft 11, 261, qytetiShkolla & " me " & FormatDateTime(Now, vbShortDate), True, 12
    End If
    objPrintClass.PrintLeft 116, 261, "Mesuesi kujdestar  ___________________", True, 12
    If Image1 <> 0 Then
        objPrintClass.PrintPicture 81, 3, Image1, 40, 40
    End If
End Sub





