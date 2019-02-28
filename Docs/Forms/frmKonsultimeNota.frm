VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{04759352-AB08-11D5-863E-0010B563AF37}#1.0#0"; "Cgridv11.ocx"
Begin VB.Form frmKonsultimeNota 
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
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
      TabIndex        =   30
      Top             =   9000
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cikli"
      ForeColor       =   &H00008000&
      Height          =   1455
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton optUlet 
         Caption         =   "Shkolla tetevjecare"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optMesme 
         Caption         =   "Shkolla e mesme"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtData 
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   19
      Top             =   9000
      Width           =   2055
   End
   Begin VB.TextBox txtIndeksi 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   840
      Width           =   735
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
      TabIndex        =   11
      Top             =   0
      Width           =   2655
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
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
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
         TabIndex        =   13
         Top             =   600
         Width           =   1095
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
         TabIndex        =   31
         Top             =   120
         Width           =   2415
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
         TabIndex        =   12
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmKonsultimeNota.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   960
      End
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9000
      Width           =   2295
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
      TabIndex        =   8
      Top             =   9000
      Width           =   2295
   End
   Begin VB.Frame fraKerkim 
      Caption         =   "Kerkim sipas ..."
      ForeColor       =   &H00008000&
      Height          =   1455
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   10095
      Begin VB.TextBox txtAtesia 
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   4680
         TabIndex        =   32
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cboVitiShkollor 
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
         ItemData        =   "frmKonsultimeNota.frx":6F75
         Left            =   8280
         List            =   "frmKonsultimeNota.frx":6F88
         TabIndex        =   15
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtAmzaNo 
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
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtKlasa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtMbiemri 
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
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtEmri 
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
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Atesia"
         Height          =   255
         Left            =   4680
         TabIndex        =   33
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Indeksi"
         Height          =   255
         Left            =   7440
         TabIndex        =   18
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblVitiShkollor 
         Caption         =   "Viti Shkollor"
         Height          =   255
         Left            =   8280
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblAmza 
         Caption         =   "Nr.Amze"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblKlasa 
         Caption         =   "Klasa"
         Height          =   255
         Left            =   6600
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblMbiemri 
         Caption         =   "Mbiemri"
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblEmri 
         Caption         =   "Emri"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
   End
   Begin Cgridv11.Cgrid Lendet 
      Height          =   945
      Left            =   0
      TabIndex        =   20
      Top             =   2760
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
   Begin Cgridv11.Cgrid Cgrid1 
      Height          =   945
      Left            =   2400
      TabIndex        =   21
      ToolTipText     =   "Klikoni mbi noten per te pare daten e marrjes "
      Top             =   2760
      Visible         =   0   'False
      Width           =   19185
      _ExtentX        =   33840
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   6960
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
   Begin VB.Label Label3 
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Lendet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblNotatDheMungesat 
      Alignment       =   2  'Center
      Caption         =   "Notat  Dhe  Mungesat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2400
      TabIndex        =   24
      Top             =   2400
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "frmKonsultimeNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VektorTeDhenash(12) As String
Dim TabelNotash(20, 3) As String
' Tabela e notave momentale
Dim TabelNotMom(20, 1) As String
Dim objPrintClass As Object
Dim objektGabimi As New clsErrorHandler
Dim dbManager As New clsDBManager
'Dim objPrintForm As New frmPrintNota
Public atesia As String
Public Vendlindja As String
Public Datelindja As String
Public Vrejtje As String
Public PrintNota As Integer
Public Rrethi As String
Public ShkEmri As String


Private Sub Cgrid1_CellClick(row As Long, col As Long)

   Dim nota             As String
   nota = Cgrid1.Text(row, col)
   If nota <> "" Then
      txtData.Text = ""
      If Not IsNull(matricaNotat(row, col)) Then
         txtData.Text = matricaNotat(row, col)

      End If
   End If
End Sub

Private Sub Cgrid1_CellLostFocus(row As Long, col As Long)
    txtData.Text = ""
End Sub

Private Sub cmdDil_Click()
   Unload Me
   Set active_form = Nothing
End Sub

Private Sub cmdOK_Click()
   
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
        Label3.Visible = True
        txtData.Visible = True
        objectInitialization KONSULTIME_NOTA_MOMENTALEI
   End If
End Sub

Private Sub cmdPrint_Click()
    
    objPrintForm.Show vbModal
    'frmPrintNota.Show
    'Set active_form = New frmPrintNota
    'Load frmPrintNota
    'frmPrintNota.SetFocus
    PrintNota = objPrintForm.nota
    Unload objPrintForm
    GetNota
    GetData
    ' *********************
    'PrintNota = 1
    Select Case PrintNota
        
        ' Nese perdoruesi zgjedh te printoje notat momentale
        Case 1
            FormatDataMom
        
        ' Nese perdoruesi zgjedh te printoje notat e semestrit te pare
        Case 2
            FormatDataSem1
        
        ' Nese perdoruesi zgjedh te printoje notat e semestrit te dyte
        Case 3
            FormatDataSem2
        
        ' Nese perdoruesi zgjedh te printoje deftesen
        Case 4
            If Me.optMesme.Value = True Then
                Set objPrintClass = CreateObject("PrintimComponent.clsFormatDeftesaMesme")
                objPrintClass.FormatDeftesaMesme VektorTeDhenash, TabelNotash
            ElseIf Me.optUlet = True Then
                Set objPrintClass = CreateObject("PrintimComponent.clsFormatDeftesaUlet")
                objPrintClass.FormatDeftesaUlet VektorTeDhenash, TabelNotash
            End If
        End Select
End Sub

Private Sub Command1_Click()
    CallHelp indeksHelp
End Sub

Private Sub Form_Activate()


  'SGGrid1.Width = Me.Width - 300
  'SGGrid1.Height = SGGrid1.Height + 900
  
End Sub

Private Sub Form_Load()
  viti
  loadForm Me
  mbushkomboboks
  Image1.Picture = LoadPicture(adresaLogo)
  lblShkolla.Caption = emerShkolla
  'inicializoGridaNotat
End Sub

Private Sub Form_Resize()
  'SGGrid1.Width = Me.Width - 300
  'fraInfo.Left = Me.Left + Me.Width - fraInfo.Width - 200
  
  'fraKerkim.Width = Me.Width - 300
  'SGGrid1.Height = SGGrid1.Height + 900
End Sub

Private Sub inicializoGridaNotat()
    
   ' gridaNotat.Width = 14575
    'gridaNotat.Rows = 2
    'gridaNotat.Cols = 3
    'gridaNotat.FixedRows = 1
    'gridaNotat.FixedCols = 0
    'gridaNotat.Row = 0
    'gridaNotat.col = 0
    'gridaNotat.ColWidth(0) = 2000
    'gridaNotat.Text = "Nr"
    'gridaNotat.col = 1
    'gridaNotat.ColWidth(1) = 2500
    'gridaNotat.Text = "Lenda"
    'gridaNotat.col = 2
    'gridaNotat.ColWidth(2) = 10000
    'gridaNotat.Text = "Notat dhe Mungesat"
    
    
End Sub

Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
End Sub

Private Sub viti()
Dim d, muai, viti As String
d = Now
Dim m, v As Integer
muai = Mid(d, 4, 2)
viti = Mid(d, 7, 4)
m = Val(muai)
v = Val(viti)
cboVitiShkollor.Text = gjej_vitin(m, v)

End Sub

' Printon te gjitha notat e grides
Public Sub FormatDataMom()
    Set objPrintim = CreateObject("PrintimComponent.clsPrintim")
    Dim fontib As Boolean
    Dim italik As Boolean
    Dim madhesi As Integer
    Dim i, j As Integer
    Dim tekst As String
    tekst = "Nxenesi " & txtEmri.Text & " " & txtMbiemri.Text _
    & " deri tani eshte vleresuar me keto nota:"
    objPrintim.PrintLeft 20, 30, tekst, False, 14
    If objPrintim.Gabim = True Then
        Exit Sub
    End If
    objPrintim.PrintLeft 20, 50, "LENDET", True, 14
    objPrintim.PrintLeft 70, 50, "NOTAT DHE MUNGESAT", True, 14
    For i = 1 To Cgrid1.Height / 255
        madhesi = 14
        ' Printon emrin e lendes korrente
        objPrintim.PrintLeft 20, 60 + (10 * i), Lendet.Text(i, 1), True, madhesi
        ' Printon notat per lenden korrente
        For j = 1 To Cgrid1.Width / 300
            ' Percaktohet nga ngjyra e cellit nese do te printohet nota si bold ose si italik
            ' Italike jane notat semestrale. Bold notat vjetore.
            If Cgrid1.CellForeColor(i, j) = vbGreen Then
                fontib = False
                italik = True
                madhesi = 12
            ElseIf Cgrid1.CellForeColor(i, j) = vbBlue Then
                fontib = True
                italik = False
                madhesi = 12
            Else
                fontib = False
                italik = False
                madhesi = 12
            End If
                objPrintim.PrintLeft 70 + (10 * j), 60 + (10 * i), Cgrid1.Text(i, j), fontib, madhesi, italik
        Next
    Next
    Printer.EndDoc
End Sub

' Printon notat e semestrit te pare, te cilat i merr nga grida
Public Sub FormatDataSem1()
    Set objPrintim = CreateObject("PrintimComponent.clsPrintim")
    Dim italik As Boolean
    Dim madhesi As Integer
    Dim i, j As Integer
    Dim tekst As String
    tekst = "Nxenesi " & txtEmri.Text & " " & txtMbiemri.Text _
    & " per semestrin e pare eshte vleresuar me keto nota:"
    objPrintim.PrintLeft 20, 30, tekst, False, 14
    If objPrintim.Gabim = True Then
        Exit Sub
    End If
    objPrintim.PrintLeft 20, 50, "LENDET", True, 14
    objPrintim.PrintLeft 70, 50, "NOTAT DHE MUNGESAT", True, 14
    
    For i = 1 To Cgrid1.Height / 255
        madhesi = 14
        ' Printon emrin e lendes korrente
        objPrintim.PrintLeft 20, 60 + (10 * i), Lendet.Text(i, 1), True, madhesi
        ' Printon notat per lenden korrente
        For j = 1 To Cgrid1.Width / 300
            ' Percaktohet nga ngjyra e cellit nese do te printohet nota si italike
            If Cgrid1.CellForeColor(i, j) = vbGreen Then
                italik = True
                madhesi = 12
                objPrintim.PrintLeft 70 + (10 * j), 60 + (10 * i), Cgrid1.Text(i, j), False, madhesi, italik
                Exit For
            ElseIf Cgrid1.CellForeColor(i, j) <> vbBlue Then
                italik = False
                madhesi = 12
                objPrintim.PrintLeft 70 + (10 * j), 60 + (10 * i), Cgrid1.Text(i, j), False, madhesi, italik
            End If
        Next
    Next
    Printer.EndDoc
End Sub

' Printon notat e semestrit te dyte, te cilat i merr nga grida
Private Sub FormatDataSem2()
    Set objPrintim = CreateObject("PrintimComponent.clsPrintim")
    Dim italik As Boolean
    Dim madhesi As Integer
    Dim i, j As Integer
    Dim tekst As String
    tekst = "Nxenesi " & txtEmri.Text & " " & txtMbiemri.Text _
    & " per semestrin e dyte eshte vleresuar me keto nota:"
    objPrintim.PrintLeft 20, 30, tekst, False, 14
    If objPrintim.Gabim = True Then
        Exit Sub
    End If
    objPrintim.PrintLeft 20, 50, "LENDET", True, 14
    objPrintim.PrintLeft 70, 50, "NOTAT DHE MUNGESAT", True, 14
    
    For i = 1 To Cgrid1.Height / 255
        madhesi = 14
        ' Printon emrin e lendes korrente
        objPrintim.PrintLeft 20, 60 + (10 * i), Lendet.Text(i, 1), True, madhesi
        For j = 1 To Cgrid1.Width / 300
            If Cgrid1.CellForeColor(i, j) = vbGreen Then
                j = j + 1
                Exit For
            End If
        Next
        
        ' Printon notat per lenden korrente
        Do While j <= Cgrid1.Width / 300
            
            ' Percaktohet nga ngjyra e cellit nese do te printohet nota si italike
            If Cgrid1.CellForeColor(i, j) = vbGreen Then
                italik = True
                madhesi = 12
                objPrintim.PrintLeft 70 + (10 * j), 60 + (10 * i), Cgrid1.Text(i, j), False, madhesi
                Exit Do
            Else
                italik = False
                madhesi = 12
                objPrintim.PrintLeft 70 + (10 * j), 60 + (10 * i), Cgrid1.Text(i, j), False, madhesi
            End If
            j = j + 1
        Loop
    Next
    Printer.EndDoc
End Sub



' Hedh ne tabelen TabNotash emrat e lendeve dhe notat semestrale dhe vjetore
' qe merren nga grida Cgrid1
Public Sub GetNota()
    Dim i, j, k As Integer
    For i = 0 To Cgrid1.Height / 255 - 1
        k = 1
        TabelNotash(i, 0) = Lendet.Text(i + 1, 1)
        For j = 1 To Cgrid1.Width / 300
            If Cgrid1.CellForeColor(i, j) = vbGreen Then
                TabelNotash(i, k) = Cgrid1.Text(i + 1, j) + MerrNote(Cgrid1.Text(i, j))
                k = k + 1
            ElseIf Cgrid1.CellForeColor(i + 1, j) = vbBlue Then
                TabelNotash(i, 3) = Cgrid1.Text(i + 1, j) + MerrNote(Cgrid1.Text(i, j))
            End If
        Next
    Next
End Sub


'Do te marre te dhenat qe duhen per te printuar deftesen dhe keto te dhena
'do i hedhe ne nje tabele sipas kesaj renditjeje:
' Rreshti 0: Emri i nxenesit
' Rreshti 1: Mbiemri i nxenesit
' Rreshti 2: Atesia
' Rreshti 3: Numri i amzes
' rreshti 4: atesia e nxenesit
' rreshti 5: emri i shkolles
' rreshti 6: vendlindja
' rreshti 7: rrethi, qyteti ku ndodhet shkolla
' rreshti 8: datelindja e nxenesit
' rreshti 9: viti shkollor
' rreshti 10: klasa
Public Sub GetData()
    VektorTeDhenash(0) = txtEmri.Text
    VektorTeDhenash(1) = txtMbiemri.Text
    VektorTeDhenash(3) = txtAmzaNo.Text
    VektorTeDhenash(5) = lblShkolla.Caption
    VektorTeDhenash(4) = atesia
    VektorTeDhenash(5) = ShkEmri
    VektorTeDhenash(6) = Vendlindja
    VektorTeDhenash(7) = Rrethi
    VektorTeDhenash(8) = Vrejtje
    VektorTeDhenash(9) = cboVitiShkollor.Text
    VektorTeDhenash(10) = txtKlasa.Text
End Sub

Public Sub GetNotaMomentale()
End Sub

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
        OrdinalInLetters = "nente"
    End Select
End Function

Private Function DigitsInLetters(numer As Integer) As String
    Dim letter As String
    Select Case numer
    Case 1
        DigitsInLetters = "nje"
    Case 2
        DigitsInLetters = "dy"
    Case 3
        DigitsInLetters = "tre"
    Case 4
        DigitsInLetters = "katër"
    Case 5
        DigitsInLetters = "pesë"
    Case 6
        DigitsInLetters = "gjashtë"
    Case 7
        DigitsInLetters = "shtatë"
    Case 8
        DigitsInLetters = "tetë"
    Case 9
        DigitsInLetters = "nëntë"
    Case 10
        DigitsInLetters = "dhjetë"
    End Select
End Function


Private Function MerrNote(tekst As String) As String
    Dim numer As Integer
    On Error GoTo Gabimi
        numer = CInt(tekst)
        MerrNote = tekst + " (" + DigitsInLetters(numer) + ")"
        Exit Function
Gabimi:
    MerrNote = ""
End Function


Private Sub mbushkomboboks()
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
End Sub

