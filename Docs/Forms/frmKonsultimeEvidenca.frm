VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{04759352-AB08-11D5-863E-0010B563AF37}#1.0#0"; "Cgridv11.ocx"
Begin VB.Form frmKonsultimeEvidenca 
   Caption         =   "Konsultime - Evidenca"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14070
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   14070
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4440
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtEmri 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   370
      Index           =   1
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   " EMRI   MBIEMRI"
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H80000009&
      Caption         =   "Printo"
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   9000
      Width           =   2055
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
      TabIndex        =   20
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton cmdWebsiste 
         BackColor       =   &H80000009&
         Caption         =   "Website"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdEmail 
         BackColor       =   &H80000009&
         Caption         =   "E-mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmKonsultimeEvidenca.frx":0000
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
         TabIndex        =   24
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblShkolla 
         Appearance      =   0  'Flat
         Caption         =   "Shkolla jo publike Nr 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   2415
      End
   End
   Begin Cgridv11.Cgrid provimet 
      Height          =   1095
      Left            =   0
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      Caption         =   "Ndihme"
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9000
      Width           =   2055
   End
   Begin VB.ComboBox cboIndeksi 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Height          =   360
      Left            =   5880
      TabIndex        =   11
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
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   7095
      Begin VB.ComboBox cboVitiShkollor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   360
         Left            =   4920
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboKlasa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   360
         Left            =   4800
         TabIndex        =   7
         Top             =   960
         Width           =   735
      End
      Begin VB.Frame fraOptions 
         Caption         =   "Nota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   3255
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
            TabIndex        =   6
            Top             =   480
            Width           =   975
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
            TabIndex        =   5
            Top             =   840
            Width           =   1695
         End
         Begin VB.OptionButton optSemestri1 
            Caption         =   "Semestale I"
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
            TabIndex        =   4
            Top             =   240
            Width           =   1695
         End
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
         TabIndex        =   12
         Top             =   720
         Width           =   735
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
         TabIndex        =   9
         Top             =   240
         Width           =   975
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
         TabIndex        =   8
         Top             =   690
         Width           =   615
      End
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9000
      Width           =   2535
   End
   Begin VB.CommandButton cmdDalje 
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9000
      Width           =   2535
   End
   Begin Cgridv11.Cgrid Emrat 
      Height          =   945
      Left            =   0
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1667
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
   Begin Cgridv11.Cgrid Cgrid2 
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      GridMode        =   0
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridBackColor   =   -2147483644
      GridForeColor   =   16777215
      GridForeColor   =   16777215
      GridBackColor   =   -2147483644
      RowHeight       =   16
      FixedColumnVisible=   -1  'True
      FixedRowVisible =   -1  'True
      ScrollBarV      =   0   'False
      ScrollBarh      =   0   'False
      Appearance3D    =   0   'False
      CellEditColor   =   0
   End
   Begin Cgridv11.Cgrid Cgrid1 
      Height          =   945
      Left            =   2400
      TabIndex        =   17
      Top             =   3480
      Visible         =   0   'False
      Width           =   19095
      _ExtentX        =   33681
      _ExtentY        =   1667
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
      CellEditColor   =   16777088
   End
   Begin Cgridv11.Cgrid Amza 
      Height          =   945
      Left            =   0
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1667
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   8040
      Visible         =   0   'False
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Arrows          =   65536
      Orientation     =   1245185
      Value           =   1
   End
End
Attribute VB_Name = "frmKonsultimeEvidenca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objektGabimi As New clsErrorHandler

Private Sub cmdDalje_Click()
   Unload Me
   Set active_form = Nothing
End Sub

Private Sub cmdDil_Click()
    CallHelp indeksHelp
End Sub

Private Sub cmdOK_Click()
   'gridaNotat.Visible = False
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      objectInitialization VISUALIZO_EVIDENCE
      If scrbar Then
 FlatScrollBar1.Visible = True
 End If
   End If
End Sub


Private Sub cmdPrint_Click()
    NoteGrida
    Dim gridHeight, gridWidth As Integer
    Dim tekst As String
    'gridWidth = Me.Cgrid2.Width * 2 / (75 * 32)
    'gridHeight = Me.Cgrid2.Height / 300
    'tekst = "||6|| 5 |7| 8 "
    'Set objPrintim = CreateObject("PrintimComponent.clsPrintim")
    'Portret = False
    'objPrintim.OrientimFaqe Portret
    
    'FormatCgrid1 tekst, 30, 30
End Sub


Private Sub Form_Load()
   loadForm Me
   viti
   mbushKomboBox
   Image1.Picture = LoadPicture(adresaLogo)
   lblShkolla.Caption = emerShkolla
   txtEmri(1).Visible = False
   'gridaNotat.Visible = False
   'objGridHandler.applyStyleGrid1 SGGrid1, "Regjistri i Notave", False
   'inicializoGridaNotat 8, 8
End Sub



Private Sub Form_Resize()
 'With SGGrid1
     ' .Width = Me.Width - 400
      
      '.Height = Me.Height
   'End With
   'Frame1.Width = Me.Width - 400
   fraInfo.Left = Me.Left + Me.Width - fraInfo.Width - 200
End Sub


Private Sub Form_Activate()
 ' With SGGrid1
   '   .Width = Me.Width - 400
      '.Height = Me.Height
  ' End With
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
    
    
    
    cboIndeksi.AddItem "A"
    cboIndeksi.AddItem "B"
    cboIndeksi.AddItem "C"
    cboIndeksi.AddItem "D"
    cboIndeksi.AddItem "E"
    cboIndeksi.AddItem "F"
    
    
    
    
    
End Sub

Private Sub viti()
Dim d, muai, viti As String
d = Now
Dim M, v As Integer
muai = Mid(d, 4, 2)
viti = Mid(d, 7, 4)
M = Val(muai)
v = Val(viti)
cboVitiShkollor.Text = gjej_vitin(M, v)

End Sub

Private Sub NoteGrida()
    Dim gridHeight, gridWidth As Integer
    Dim i, j, k As Integer
    Dim tekst As String
    Dim Portret As Boolean
    gridWidth = Me.Cgrid2.Width * 2 / (75 * 32)
    gridHeight = Me.Cgrid1.Height / 255
    tekst = "||6|| 5 |7| 8 "
    Set objPrintim = CreateObject("PrintimComponent.clsPrintim")
    Portret = False
    objPrintim.OrientimFaqe Portret
    
    objPrintim.PrintLeft 10, 10, "Emri Mbiemri", True, 14
    ' Nese perdoruesi anullon printimin
    If objPrintim.Gabim = True Then
        Exit Sub
    End If
    
    ' Printo emrat e nxenesve
    For i = 1 To gridWidth
        objPrintim.PrintLeft 60 + (30 * (i - 1)), 30, Cgrid2.Text(1, i), True, 12
        k = 1
        For j = 1 To gridHeight
            objPrintim.PrintLeft 10, 20 + (10 * k), Emrat.Text(i, 1), True, 10
            FormatCgrid1 Cgrid1.Text(i, j), 60 + (30 * (i - 1)), 20 + (10 * j)
            k = k + 1
            If (20 + (10 * j)) > 190 Then
                Printer.NewPage
                k = 1
            End If
            'Printer.Print
        Next
    Next
        
    ' Printo emrat e nxenesve
    For j = 1 To gridWidth
        For i = 1 To gridHeight
            objPrintim.PrintLeft 10, 20 + (10 * i), Emrat.Text(i, 1), True, 10
            'if 20 + (10 * i)
            'objPrintim.PrintLeft 40 + (30 * i), 20 + (10 * j), Cgrid1.Text(i, j)
        Next
    Next
    Printer.EndDoc
End Sub

' Metode qe ben formatimin dhe printimin e nje qelize te grides se notave
Private Sub FormatCgrid1(Qeliza As String, PosX As Long, PosY As Long)
    Dim i, j, k As Integer
    i = 1
    j = 0
    Do While i < Len(Qeliza)
        If Mid(Qeliza, i, 2) = "||" Then
            If Mid(Qeliza, i + 1, 1) <> "|" Then
                k = 2
            Else
                k = 1
            End If
            objPrintim.PrintLeft PosX + (7 * j), PosY, Mid(Qeliza, i + 2, k), True, 12
            i = i + k + 4
            j = j + 1
        ElseIf Mid(Qeliza, i, 1) = "|" Then
            objPrintim.PrintLeft PosX + (7 * j), PosY, Mid(Qeliza, i + 1, 1), True, 10
            i = i + 3
            j = j + 1
        ElseIf Mid(Qeliza, i, 1) = " " Then
            i = i + 1
        Else
            objPrintim.PrintLeft PosX + (7 * j), PosY, Mid(Qeliza, i, 1), False, 10
            i = i + 1
            j = j + 1
        End If
    Loop
    Printer.EndDoc
End Sub

