VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmHedhjeNota 
   Caption         =   "Hedhje Te Dhenash - Nota"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   13530
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridaNotat 
      Height          =   5175
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9128
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Frame fraInfo 
      Height          =   1455
      Left            =   10560
      TabIndex        =   17
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdWebsiste 
         Caption         =   "Website"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "E-mail"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   1080
         Left            =   1320
         Picture         =   "frmHedhjeNota.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label lblWebsite 
         Height          =   375
         Left            =   1560
         TabIndex        =   20
         Top             =   2040
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdDalje 
      BackColor       =   &H00FFFFFF&
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7800
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.ComboBox cboVitiShkollor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8160
         TabIndex        =   16
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox cboIndeksi 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5760
         TabIndex        =   14
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cboKlasa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Format          =   68419585
         CurrentDate     =   38313
      End
      Begin VB.Frame fraOptions 
         Caption         =   "Nota"
         Height          =   1095
         Left            =   2880
         TabIndex        =   1
         Top             =   120
         Width           =   6735
         Begin VB.OptionButton optDetyreKontrolli 
            Caption         =   "Detyre Kontrolli"
            Height          =   255
            Left            =   2280
            TabIndex        =   9
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optMomentaleSII 
            Caption         =   "Momentale Semestri II"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton optMomentaleSI 
            Caption         =   "Momentale Semestri I "
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton optSemestriI 
            Caption         =   "Semestale I"
            Height          =   255
            Left            =   4080
            TabIndex        =   4
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optSemestraleII 
            Caption         =   "Semestrale II"
            Height          =   255
            Left            =   4080
            TabIndex        =   3
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton optVjetore 
            Caption         =   "Vjetore"
            Height          =   255
            Left            =   5520
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label lblVitiShkollor 
         Caption         =   "Viti Shkollor"
         Height          =   255
         Left            =   7200
         TabIndex        =   15
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label lblIndeksi 
         Caption         =   "Indeksi"
         Height          =   255
         Left            =   5040
         TabIndex        =   13
         Top             =   1350
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Klasa"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   1350
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmHedhjeNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdDalje_Click()
  Unload Me
  Set active_form = Nothing
End Sub

Private Sub cmdDil_Click()
    CallHelp indeksHelp
End Sub

Private Sub cmdOK_Click()
  
  '' to do
  Unload Me
End Sub



Private Sub Form_Activate()
 ' With SGGrid1
  '    .Width = Me.Width - 400
      '.Height = Me.Height
  ' End With
  
End Sub

Private Sub Form_Load()

  loadForm Me
  
  'MSFlexGrid1.AddItem 1
  
  

  'SGGrid1.Columns.RemoveAll False
  
'  With objGridHandler
 '    .applyStyleGrid1 SGGrid1, "Hedhja e Notave Te Nxenesve", True
     
     
  
    ' .mergeGrid SGGrid1
  
  
 ' End With
 
 inicializoGridaNotat
 viti
  
End Sub

Private Sub Form_Resize()
' With SGGrid1
 '     .Width = Me.Width - 400
      '.Height = Me.Height
 '  End With
   
   fraInfo.Left = Me.Left + Me.Width - fraInfo.Width - 200
End Sub



Private Sub inicializoGridaNotat()
    
    gridaNotat.Width = 14575
    gridaNotat.Rows = 2
    gridaNotat.Cols = 3
    gridaNotat.FixedRows = 1
    gridaNotat.FixedCols = 0
    gridaNotat.row = 0
    gridaNotat.col = 0
    gridaNotat.ColWidth(0) = 2000
    gridaNotat.Text = "Nr"
    gridaNotat.col = 1
    gridaNotat.ColWidth(1) = 2500
    gridaNotat.Text = "Lenda"
    gridaNotat.col = 2
    gridaNotat.ColWidth(2) = 10000
    gridaNotat.Text = "Notat dhe Mungesat"
    
    
End Sub

Private Sub viti()
   Dim d, muai, viti    As String
   d = Now
   Dim m, v             As Integer
   muai = Mid(d, 4, 2)
   viti = Mid(d, 7, 4)
   m = Val(muai)
   v = Val(viti)
   cboVitiShkollor.Text = gjej_vitin(m, v)

End Sub
