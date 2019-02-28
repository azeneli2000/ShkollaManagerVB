VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmShenime 
   Caption         =   "Shenime per sjelljen"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   15240
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtbSjellja 
      Height          =   4095
      Left            =   7560
      TabIndex        =   20
      Top             =   2880
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7223
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmShenime.frx":0000
   End
   Begin VB.CommandButton cmdNdihme 
      BackColor       =   &H00FFFFFF&
      DisabledPicture =   "frmShenime.frx":0082
      DownPicture     =   "frmShenime.frx":35BC
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      Picture         =   "frmShenime.frx":6AF6
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8640
      Width           =   2535
   End
   Begin VB.Frame fraInfo 
      Height          =   1455
      Left            =   12360
      TabIndex        =   14
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmShenime.frx":A030
         Stretch         =   -1  'True
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblWebsite 
         Height          =   375
         Left            =   1560
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmShenime.frx":10FA5
      DownPicture     =   "frmShenime.frx":173E7
      Height          =   375
      Left            =   4080
      Picture         =   "frmShenime.frx":1D829
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8640
      Width           =   2535
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmShenime.frx":23C6B
      DownPicture     =   "frmShenime.frx":2A0AD
      Height          =   375
      Left            =   7440
      Picture         =   "frmShenime.frx":304EF
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8640
      Width           =   2535
   End
   Begin VB.ListBox listAmza 
      Appearance      =   0  'Flat
      Height          =   4905
      Left            =   5640
      TabIndex        =   9
      Top             =   2880
      Width           =   1575
   End
   Begin VB.ListBox lista 
      Appearance      =   0  'Flat
      Height          =   4905
      Left            =   360
      TabIndex        =   8
      Top             =   2880
      Width           =   4095
   End
   Begin VB.Frame fraKerkimi 
      Caption         =   "Kerko sipas"
      ForeColor       =   &H00008000&
      Height          =   1935
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      Begin VB.CommandButton cmdKerko 
         BackColor       =   &H80000009&
         DisabledPicture =   "frmShenime.frx":36931
         DownPicture     =   "frmShenime.frx":3CD73
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         Picture         =   "frmShenime.frx":431B5
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox cboIndeksi 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2000
      End
      Begin VB.ComboBox cboKlasa 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2000
      End
      Begin VB.ComboBox cboVitiShkollor 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   2000
      End
      Begin VB.Label Label3 
         Caption         =   "Indeksi"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Klasa"
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Viti shkollor"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sjellja e nxenesit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   11
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Emrat e nxenesve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   4095
   End
End
Attribute VB_Name = "frmShenime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim objUIController As clsUIController
Dim objektGabimi As New clsErrorHandler

Private Sub cmdDil_Click()
   Unload Me
   Set active_form = Nothing
   
   
End Sub

Private Sub cmdEmail_Click()
    If Not OpenEmailProgram(email) Then
        MsgBox "Ju nuk keni asnje program te instaluar per te derguar email", vbExclamation, "Gabim ne dergimin e emailit"
    End If
End Sub

Private Sub cmdKerko_Click()
    rtbSjellja.Text = ""
    lista.Clear
    objectInitialization HEDHJE_SJELLJE_KERKO
End Sub

Private Sub cmdNdihme_Click()
CallHelp indeksHelp
End Sub

Private Sub cmdOK_Click()
   Dim data             As String
   data = date
   objectInitialization HEDHJA_SJELLJA_OK
   If data_modifikimi = "" Or data <> data_modifikimi Then
      objectInitialization MODIFIKIMI_I_DATES
   End If
   rtbSjellja.Text = ""

End Sub



Private Sub cmdWebsiste_Click()
    If website <> "" Then
        GoToWeb website
    Else
        MsgBox "Ju nuk e keni dhene adresen e faqes tuaj te web-it.", vbInformation
    End If
End Sub

Private Sub Form_Load()
  loadForm Me
  Image1.Picture = LoadPicture(adresaLogo)
  lblShkolla.Caption = emerShkolla
  mbushKomboBox
  viti
  listAmza.Visible = False
  lista.ListIndex = -1
  percaktoTeDrejtat
  cmdOK.Enabled = False
End Sub
Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
End Sub
Private Sub Form_Resize()

   
  ' fraInfo.Left = Me.Left + Me.Width - fraInfo.Width - 200
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
   Dim d, muai, viti    As String
   d = Now
   Dim M, v             As Integer
   muai = DateTime.Month(DateTime.Now)
   viti = DateTime.Year(DateTime.date)
   M = Val(muai)
   v = Val(viti)
   cboVitiShkollor.Text = gjej_vitin(M, v)

End Sub

Private Sub percaktoTeDrejtat()
    
    Select Case statusi
        
        Case "SupervizorEmesme"
            mbushKomboMesme
        
        Case "SupervizorTetevjecare"
            mbushKomboTetevjecare
        Case Else
            
    End Select
    
            
            
End Sub

Private Sub mbushKomboMesme()
    cboKlasa.Clear
    cboKlasa.AddItem "9"
    cboKlasa.AddItem "10"
    cboKlasa.AddItem "11"
    cboKlasa.AddItem "12"
End Sub

Private Sub mbushKomboTetevjecare()
   cboKlasa.Clear
   cboKlasa.AddItem "1"
   cboKlasa.AddItem "2"
   cboKlasa.AddItem "3"
   cboKlasa.AddItem "4"
   cboKlasa.AddItem "5"
   cboKlasa.AddItem "6"
   cboKlasa.AddItem "7"
   cboKlasa.AddItem "8"
End Sub

