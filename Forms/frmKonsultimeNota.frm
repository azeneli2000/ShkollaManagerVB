VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{04759352-AB08-11D5-863E-0010B563AF37}#1.0#0"; "Cgridv11.ocx"
Begin VB.Form frmKonsultimeNota 
   Caption         =   "Konsultimi i notave te nxenesit"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox lblLargimi 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox lblR 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "R"
      Top             =   3050
      Width           =   255
   End
   Begin VB.TextBox lblV 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   43
      Text            =   "V"
      Top             =   3050
      Width           =   255
   End
   Begin VB.TextBox lblS2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "S2"
      Top             =   3050
      Width           =   255
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
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Label3 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "Data"
      Top             =   2240
      Width           =   1215
   End
   Begin VB.TextBox lblS1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   41
      Text            =   "S1"
      Top             =   3050
      Width           =   255
   End
   Begin VB.TextBox lblNotatDheMungesat 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   300
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "Notat dhe Mungesat"
      Top             =   2950
      Width           =   2655
   End
   Begin VB.TextBox label2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "Lendet"
      Top             =   2950
      Width           =   1215
   End
   Begin MSComCtl2.FlatScrollBar fsbVertical 
      Height          =   5415
      Left            =   15000
      TabIndex        =   36
      Top             =   3360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   9551
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1245184
   End
   Begin MSComCtl2.FlatScrollBar fsbHorizontali 
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   4800
      Visible         =   0   'False
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Arrows          =   65536
      Orientation     =   1245185
      Value           =   1
   End
   Begin VB.Frame fraOpsioni 
      Caption         =   "Lloji i notave :"
      ForeColor       =   &H00008000&
      Height          =   1935
      Left            =   1800
      TabIndex        =   30
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton optPerfundimtare 
         Caption         =   "Perfundimtare"
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   3255
      End
      Begin VB.OptionButton optSemestri2 
         Caption         =   "Semestri II"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   2175
      End
      Begin VB.OptionButton optSemestri1 
         Caption         =   "Semestri I"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmKonsultimeNota.frx":0000
      DownPicture     =   "frmKonsultimeNota.frx":353A
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      Picture         =   "frmKonsultimeNota.frx":6A74
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9120
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cikli"
      ForeColor       =   &H00008000&
      Height          =   1935
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton optUlet 
         Caption         =   "Shkolla nëntëvjeçare"
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optMesme 
         Caption         =   "Shkolla e mesme"
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmKonsultimeNota.frx":9FAE
      DownPicture     =   "frmKonsultimeNota.frx":D4E8
      Height          =   375
      Left            =   12360
      Picture         =   "frmKonsultimeNota.frx":10A22
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9120
      Width           =   2535
   End
   Begin VB.TextBox txtIndeksi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1440
      Width           =   735
   End
   Begin VB.Frame fraInfo 
      Height          =   1575
      Left            =   12480
      TabIndex        =   10
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   600
         Width           =   1095
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
         TabIndex        =   27
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblWebsite 
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmKonsultimeNota.frx":13F5C
         Stretch         =   -1  'True
         Top             =   600
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmKonsultimeNota.frx":1AED1
      DownPicture     =   "frmKonsultimeNota.frx":21313
      Height          =   375
      Left            =   8880
      Picture         =   "frmKonsultimeNota.frx":27755
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9120
      Width           =   2535
   End
   Begin VB.Frame fraKerkim 
      Caption         =   "Kerkim sipas ..."
      ForeColor       =   &H00008000&
      Height          =   1935
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton cmdKerko 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "frmKonsultimeNota.frx":2DB97
         DownPicture     =   "frmKonsultimeNota.frx":33FD9
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
         Left            =   4560
         Picture         =   "frmKonsultimeNota.frx":3A41B
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtAtesia 
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
         Height          =   375
         Left            =   4680
         TabIndex        =   28
         Top             =   600
         Width           =   1695
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
         Height          =   360
         ItemData        =   "frmKonsultimeNota.frx":4085D
         Left            =   1800
         List            =   "frmKonsultimeNota.frx":40870
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtAmzaNo 
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
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtKlasa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtMbiemri 
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
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtEmri 
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
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Atesia"
         Height          =   255
         Left            =   4680
         TabIndex        =   29
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Indeksi"
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblVitiShkollor 
         Caption         =   "Viti Shkollor"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblAmza 
         Caption         =   "Nr.Amze"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblKlasa 
         Caption         =   "Klasa"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblMbiemri 
         Caption         =   "Mbiemri"
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblEmri 
         Caption         =   "Emri"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      MousePointer    =   1  'Arrow
      TabIndex        =   35
      Top             =   8715
      Width           =   15300
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   3360
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   15500
   End
   Begin Cgridv11.Cgrid gridaProvime 
      Height          =   2175
      Left            =   240
      TabIndex        =   44
      Top             =   3360
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3836
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
   Begin Cgridv11.Cgrid Lendet 
      Height          =   1305
      Left            =   0
      TabIndex        =   45
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2302
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
      Height          =   2625
      Left            =   2400
      TabIndex        =   46
      Top             =   3360
      Visible         =   0   'False
      Width           =   19185
      _ExtentX        =   33840
      _ExtentY        =   4630
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
End
Attribute VB_Name = "frmKonsultimeNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VektorTeDhenash(15) As String
Dim TabelNotash(25, 4) As String
' Tabela e notave momentale
Dim TabelNotMom(20, 1) As String
Dim objPrintClass As Object
Dim objektGabimi As New clsErrorHandler
Dim dbManager As New clsDBManager
Dim objPrintForm As New frmPrintNota
Dim gabimPrintim As Boolean
Public atesia As String
Public Vendlindja As String
Public Datelindja As String
Public Vrejtje As String
Public PrintNota As Integer
'Public Rrethi As String
'Public ShkEmri As String
'Public Adresa As String
'Public Qyteti As String
Public MungesaS1Me, MungesaS2Me, MungesaS1Pa, MungesaS2Pa As Integer


Private Sub Cgrid1_CellClick(row As Long, col As Long)
   Dim nota             As String
   nota = Cgrid1.Text(row, col)
   txtData.Text = ""
   If nota <> "" Then
      
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

Private Sub cmdEmail_Click()
    If Not OpenEmailProgram(email) Then
        MsgBox "Ju nuk keni asnje program te instaluar per te derguar email", vbExclamation, "Gabim ne dergimin e emailit"
    End If
End Sub

Private Sub cmdKerko_Click()
   cmdKerko.Enabled = False
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      Label3.Visible = True
      txtData.Visible = True
      objectInitialization KONTROLLI_I_NOTAVE       ' KONSULTIME_NOTA_MOMENTALEI
   End If
   cmdKerko.Enabled = True
End Sub



Private Sub cmdPrint_Click()
   If (optUlet Or optMesme) And optSemestri1 Then
        FormatDataSem 1
   ElseIf (optUlet Or optMesme) And optSemestri2 Then
        FormatDataSem 2
   ElseIf optPerfundimtare Then
        objPrintForm.show vbModal
        PrintNota = objPrintForm.nota
        If PrintNota = 1 Then
            GetData
            GetNota
            FormatNotaPerf
        ElseIf PrintNota = 2 Then
            GetData
            GetNota
            PrintNotaPerfSem 1
        ElseIf PrintNota = 3 Then
            GetData
            GetNota
            PrintNotaPerfSem 2
        ElseIf PrintNota = 4 Then
            GetData
            GetNota
            Dim nResult As Integer
            Set objPrintClass = CreateObject("PrintimComponent.clsFormatDeftesatSeri")
            ' Kontrollon nese ka printer te instaluar
            If objPrintClass.PrinterIsInstalled = False Then
                Exit Sub
            End If
            ' Nese nxenesi eshte ne shkollen e mesme
            If Me.optMesme.Value = True Then
                nResult = MsgBox("Futeni deftesen ne printer ne menyre qe te printohen gjeneralitet.", vbOKCancel + vbInformation, "Printimi i defteses")
                If nResult = vbOK Then
                    'objPrintClass.OrientimFaqe False
                    objPrintClass.vitishkollor = cboVitiShkollor.Text
                    If objPrintClass.Gabim = True Then
                        MsgBox "Printeri juaj nuk lejon qe formati i printerit tuaj te nderrohet ne " _
                            & "menyre automatike nga Portrait ne Landscape. Ju lutemi nderrojeni formatin " _
                            & "ne menyre manuale tek preferencat e printerit tuaj.", vbOKOnly + vbCritical, "Gabim ne printim"
                        Exit Sub
                    End If
                    objPrintClass.PaperSizeA4 False
                    If objPrintClass.Gabim = True Then
                        MsgBox "Printeri juaj nuk mund te printoje dot ne format A3, i " _
                            & "nevojshem per printimin e deftesave te shkolles se mesme", _
                            vbOKOnly + vbCritical, "Gabim ne printim"
                    End If
                    objPrintClass.DeftesaMesmeSeriFPa VektorTeDhenash
                    If objPrintClass.Gabim = True Then
                        Exit Sub
                    End If
                    objPrintClass.EndDoc
                    Else
                        Exit Sub
                    End If
                    nResult = MsgBox("Ju lutemi kthejeni deftesen nga ana tjeter ne printer", vbOKCancel + vbInformation, "Duke printuar...")
                    If nResult = vbOK Then
                        objPrintClass.OrientimFaqe False
                        If objPrintClass.Gabim = True Then
                            MsgBox "Printeri juaj nuk lejon qe formati i printerit tuaj te nderrohet ne " _
                                & "menyre automatike nga Portrait ne Landscape. Ju lutemi nderrojeni formatin " _
                                & "ne menyre manuale tek preferencat e printerit tuaj.", vbOKOnly + vbCritical, "Gabim ne printim"
                            Exit Sub
                        End If
                        objPrintClass.PaperSizeA4 False
                    If objPrintClass.Gabim = True Then
                        MsgBox "Printeri juaj nuk mund te printoje dot ne format A3, i " _
                            & "nevojshem per printimin e deftesave te shkolles se mesme", _
                            vbOKOnly + vbCritical, "Gabim ne printim"
                    End If
                        objPrintClass.DeftesaMesmeSeriFPr TabelNotash, VektorTeDhenash
                        If objPrintClass.Gabim = True Then
                            Exit Sub
                        End If
                        objPrintClass.EndDoc
                Else
                    Exit Sub
                End If
                
                ' Nese nxenesi eshte ne shkollen tetevjecare
            ElseIf Me.optUlet = True Then
                nResult = MsgBox("Futeni deftesen ne printer ne menyre qe te printohen gjeneralitet.", vbOKCancel + vbInformation, "Printimi i defteses")
                If nResult = vbOK Then
                    objPrintClass.OrientimFaqe True
                    objPrintClass.PaperSizeA4 True
                    objPrintClass.vitishkollor = cboVitiShkollor.Text
                    objPrintClass.DeftesaUletSeriFPa VektorTeDhenash
                    ' Nese perdoruesi anullon printimin
                    If objPrintClass.Gabim = True Then
                        Exit Sub
                    End If
                    objPrintClass.EndDoc
                Else
                    Exit Sub
                End If
                nResult = MsgBox("Ju lutemi kthejeni deftesen nga ana tjeter ne printer", vbOKCancel + vbInformation, "Duke printuar...")
                If nResult = vbOK Then
                    objPrintClass.DeftesaUletSeriFPr TabelNotash, VektorTeDhenash
                    ' Nese perdoruesi anullon printimin
                    If objPrintClass.Gabim = True Then
                        Exit Sub
                    End If
                    objPrintClass.EndDoc
                Else
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub


Private Sub cmdWebsiste_Click()
    If website <> "" Then
        GoToWeb website
    Else
        MsgBox "Ju nuk e keni dhene adresen e faqes tuaj te web-it.", vbInformation
    End If
End Sub

Private Sub Command1_Click()
    CallHelp indeksHelp
End Sub





Private Sub FlatScrollBar1_Change()
    
End Sub

Private Sub Form_Activate()


  'SGGrid1.Width = Me.Width - 300
  'SGGrid1.Height = SGGrid1.Height + 900
 
End Sub

Private Sub Form_Load()
  cboVitiShkollor.Clear
  mbushkomboboks
  viti
  loadForm Me
  Image1.Picture = LoadPicture(adresaLogo)
  lblShkolla.Caption = emerShkolla
  percaktoRendinSipasTabit
  fsbVertical.Visible = False
  Label2.Visible = False
  Label3.Visible = False
  lblS1.Visible = False
  lblS2.Visible = False
  lblR.Visible = False
  lblV.Visible = False
  lblNotatDheMungesat.Visible = False
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
   Dim d, muai, viti    As String
   d = Now
   Dim M, v             As Integer
   muai = DateTime.Month(DateTime.Now)
   viti = DateTime.Year(DateTime.date)
   M = Val(muai)
   v = Val(viti)
   cboVitiShkollor.Text = gjej_vitin(M, v)

End Sub


' Printon notat e semestrit te zgjedhur, te cilat i merr nga grida
Private Sub FormatDataSem(semestri As Integer)
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
    If semestri = 1 Then
        semesterString = "parë"
    Else
        semesterString = "dytë"
    End If
    tekst = "Nxënësi " & txtEmri.Text & " " & txtMbiemri.Text & " ne klasen '" & txtKlasa.Text & " " & txtIndeksi.Text _
    & "' për semestrin e " & semesterString & " është vlerësuar me këto nota:"
    objPrintClass.PrintLeft 20, 88, tekst, False, 12
    If objPrintClass.Gabim = True Then
        Exit Sub
    End If
    PrintoFormatPergj
    'I = Len(lblShkolla.Caption)
    'objPrintClass.PrintLeft 276 - I * 1.5, 10, "Shkolla " & lblShkolla, True, 10, False
    objPrintClass.PrintTabele 30, 105, 11, 6, 21, 1
    objPrintClass.PrintTabele 41, 105, 57, 6, 21, 1
    objPrintClass.PrintTabele 98, 105, 100, 6, 21, 1
    For i = 1 To 20
        objPrintClass.PrintLeft 31, 106.5 + (6 * i), CStr(i), False, 12, False
    Next
    objPrintClass.PrintLeft 31, 106, "Nr", True, 12
    objPrintClass.PrintLeft 45, 106, "Lendet", True, 12
    objPrintClass.PrintLeft 101, 106, "Notat", True, 12
    mungArs = 0
    mungPaArs = 0
    a = 1
    For i = 1 To Cgrid1.Height / 255
        b = 1
        ' Printon emrin e lendes korrente
        objPrintClass.PrintLeft 42, 106 + (6 * a), Lendet.Text(i, 1), False, 12
        ' Printon notat per lenden korrente
        For j = 1 To Cgrid1.Width / 300
            ' Percaktohet nga ngjyra e cellit nese do te printohet nota si italike
            If Cgrid1.Text(i, j) = " m" Then
                If Cgrid1.CellForeColor(i, j) = vbRed Then
                    mungArs = mungArs + 1
                Else
                    mungPaArs = mungPaArs + 1
                End If
            ElseIf Cgrid1.CellForeColor(i, j) = &HC00000 Then
                Exit For
            ElseIf Cgrid1.CellForeColor(i, j) <> &HFF& Then
                italik = False
                objPrintClass.PrintLeft 93 + (6 * b), 106.5 + (6 * a), Cgrid1.Text(i, j), False, 12, False
                b = b + 1
            End If
        Next
        a = a + 1
    Next
    objPrintClass.PrintLeft 86, 238, "MUNGESA", True, 12, False
    objPrintClass.PrintLeft 80, 244, "me arsye     pa arsye", False, 12, False
    objPrintClass.PrintLeft 89, 250, CStr(mungArs) & Space(16) & CStr(mungPaArs)
    objPrintClass.EndDoc
End Sub

' Printon notat perfundimtare semestrale dhe vjetore te nxenesit
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
    objPrintClass.PrintLeft 90, 63, "VERTETIM", True, 12, False, True
    If objPrintClass.Gabim = True Then
        Exit Sub
    End If
    PrintoFormatPergj
    tekst = "Vërtetojmë se nxënësi " & txtEmri.Text & " " & txtMbiemri.Text & _
        " me numër amze " & txtAmzaNo.Text & " në vitin shkollor " & cboVitiShkollor.Text
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
    Do While TabelNotash(i - 1, 0) <> "" And i <= 20
        objPrintClass.PrintLeft 57, 114.5 + (6 * i), TabelNotash(i - 1, 0), False, 12, False
        objPrintClass.PrintLeft 117, 114.5 + (6 * i), TabelNotash(i - 1, 1), False, 12, False
        objPrintClass.PrintLeft 136, 114.5 + (6 * i), TabelNotash(i - 1, 2), False, 12, False
        objPrintClass.PrintLeft 156, 114.5 + (6 * i), TabelNotash(i - 1, 3), False, 12, False
        i = i + 1
    Loop
    objPrintClass.EndDoc
End Sub

Private Sub PrintoFormatPergj()
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
    If optPerfundimtare.Value = True Then
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

Private Sub PrintNotaPerfSem(semestri As Integer)
    'If GetObject(objPrintClass) = Nothing Then
    'End If
    'Dim obj As Object
    'obj = GetObject(objPrintClass)
    If (objPrintClass Is Nothing) Then
        Set objPrintClass = CreateObject("PrintimComponent.clsPrintim")
    End If
    If objPrintClass.PrinterIsInstalled = False Then
        Exit Sub
    End If
    Dim tekst As String
    Dim semesterString As String
    Dim mungesaMe, mungesaPa As Integer
    Dim i, j As String
    objPrintClass.PrintFont "Times New Roman"
    objPrintClass.OrientimFaqe True
    If objPrintClass.Gabim = True Then
        MsgBox "Printeri juaj nuk lejon qe te nderrohet ne menyre automatike formati ne Portrait. " & _
            "Nderrojeni formatin nga Landscape ne Portrait te preferencat e printerit tuaj.", vbCritical + vbOKOnly, _
            "Gabim ne Printim."
        Exit Sub
    End If
    If semestri = 1 Then
        semesterString = "PARE"
        mungesaMe = MungesaS1Me
        mungesaPa = MungesaS1Pa
    Else
        semesterString = "DYTE"
        mungesaMe = MungesaS2Me
        mungesaPa = MungesaS2Pa
    End If
    tekst = "PERFUNDIMET E SEMESTRIT TE " & semesterString
    objPrintClass.PrintLeft 60, 70, tekst, True, 12, False, True
    If objPrintClass.Gabim = True Then
        Exit Sub
    End If
    PrintoFormatPergj
    objPrintClass.PrintLeft 73, 82, "Nxënësi   " & txtEmri.Text & " " & txtMbiemri.Text, False, 12, False, True
    objPrintClass.PrintLeft 60, 88, "Klasa " & txtKlasa.Text & " " & txtIndeksi.Text & _
    "   Viti Shkollor " & cboVitiShkollor.Text & "  Nr i amzës " & txtAmzaNo.Text
    objPrintClass.PrintTabele 63, 105, 11, 6, 21, 1
    objPrintClass.PrintTabele 74, 105, 57, 6, 21, 1
    objPrintClass.PrintTabele 131, 105, 18, 6, 21, 1
    objPrintClass.PrintLeft 64, 106.5, "Nr", True, 12
    objPrintClass.PrintLeft 90, 106.5, "Lendet", True, 12
    objPrintClass.PrintLeft 135, 106.5, "Nota", True, 12
    For i = 1 To 20
        objPrintClass.PrintLeft 65, 106.5 + (6 * i), CStr(i), False, 12
    Next
    i = 1
    Do While TabelNotash(i - 1, 0) <> ""
        objPrintClass.PrintLeft 75, 106.5 + (6 * i), TabelNotash(i - 1, 0), False, 12, False
        objPrintClass.PrintLeft 134, 106.5 + (6 * i), TabelNotash(i - 1, semestri), False, 12, True
        ' Pasi bejme printimin, fshijme nje nga nje elementet e tabeles TabelNotash
        TabelNotash(i - 1, 0) = ""
        TabelNotash(i - 1, 1) = ""
        TabelNotash(i - 1, 2) = ""
        TabelNotash(i - 1, 3) = ""

        i = i + 1
    Loop
    objPrintClass.PrintLeft 86, 238, "MUNGESA", True, 12, False
    objPrintClass.PrintLeft 80, 244, "me arsye     pa arsye", False, 12, False
    objPrintClass.PrintLeft 89, 250, CStr(mungesaMe) & Space(16) & CStr(mungesaPa)
    objPrintClass.EndDoc
    
End Sub

' Hedh ne tabelen TabNotash emrat e lendeve dhe notat semestrale dhe vjetore
' qe merren nga grida Cgrid1
Public Sub GetNota()
    Dim i, j, k As Integer
    For i = 0 To Cgrid1.Height / 255 - 1
        k = 1
        TabelNotash(i, 0) = Lendet.Text(i + 1, 1)
        'For j = 1 To Cgrid1.Width / 300
            TabelNotash(i, 1) = Cgrid1.Text(i + 1, 1)
            TabelNotash(i, 2) = Cgrid1.Text(i + 1, 2)
            TabelNotash(i, 3) = Cgrid1.Text(i + 1, 3)
            If (Cgrid1.Width / 300) = 4 Then
                TabelNotash(i, 4) = Cgrid1.Text(i + 1, 4)
            End If
            If Cgrid1.CellForeColor(i + 1, j) = &HC00000 Then
                TabelNotash(i, j) = Cgrid1.Text(i + 1, j)
                'k = k + 1
            ElseIf Cgrid1.CellForeColor(i + 1, j) = &HFF& Then
                TabelNotash(i, 3) = Cgrid1.Text(i + 1, j)
            End If
        'Next
    Next
End Sub


'Do te marre te dhenat qe duhen per te printuar deftesen dhe keto te dhena
'do i hedhe ne nje tabele sipas kesaj renditjeje:
' Rreshti 0: Emri i nxenesit
' Rreshti 1: Mbiemri i nxenesit
' Rreshti 2: Atesia
' Rreshti 3: Numri i amzes
' Rreshti 4: Vrejtje
' rreshti 5: emri i shkolles
' rreshti 6: vendlindja
' rreshti 7: rrethi, qyteti ku ndodhet shkolla
' rreshti 8: datelindja e nxenesit
' rreshti 9: viti shkollor
' rreshti 10: klasa
Public Sub GetData()
    VektorTeDhenash(0) = txtEmri.Text
    VektorTeDhenash(1) = txtMbiemri.Text
    VektorTeDhenash(2) = txtAtesia.Text
    VektorTeDhenash(3) = txtAmzaNo.Text
    VektorTeDhenash(4) = Vrejtje
    VektorTeDhenash(5) = lblShkolla.Caption
    VektorTeDhenash(6) = Vendlindja
    VektorTeDhenash(7) = rrethiShkolla
    VektorTeDhenash(8) = Datelindja
    VektorTeDhenash(9) = cboVitiShkollor.Text
    VektorTeDhenash(10) = txtKlasa.Text
    
    objectInitialization MUNGESAT_ME_PA
    VektorTeDhenash(11) = CStr(MungesaS1Me)
    VektorTeDhenash(12) = CStr(MungesaS2Me)
    VektorTeDhenash(13) = CStr(MungesaS1Pa)
    VektorTeDhenash(14) = CStr(MungesaS2Pa)
End Sub


Public Sub GetNotaMomentale()
End Sub


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
    'cboVitiShkollor.Clear
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
End Sub

Private Sub percaktoRendinSipasTabit()
    txtAmzaNo.TabIndex = 0
    txtEmri.TabIndex = 1
    txtMbiemri.TabIndex = 2
    txtAtesia.TabIndex = 3
    
End Sub

Private Sub percaktoTeDrejtat()

    Select Case statusi
        Case "SupervizorEmesme"
                
            Me.optMesme.Visible = False
            Me.optUlet.Visible = False
            Me.optMesme.Value = True
            Me.optUlet.Value = False
            
        Case "SupervizorTetevjecare"
        
            Me.optMesme.Visible = False
            Me.optUlet.Visible = False
            Me.optMesme.Value = False
            Me.optUlet.Value = True
            
        Case "Vizitor"
        
        
    End Select
            
End Sub




Private Sub fsbHorizontali_Change()
    Cgrid1.left = 2400 - fsbHorizontali.Value
End Sub

Private Sub fsbHorizontali_Scroll()
    Cgrid1.left = 2400 - fsbHorizontali.Value
End Sub

Private Sub fsbVertikali_Change()
    Cgrid1.top = 3360 - fsbVertikali.Value
    Lendet.top = 3360 - fsbVertikali.Value
End Sub

Private Sub fsbVertikali_Scroll()
   Cgrid1.top = 3360 - fsbVertikali.Value
   Lendet.top = 3360 - fsbVertikali.Value
End Sub

Private Sub fsbVertical_Change()
    Cgrid1.top = 3360 - fsbVertical.Value
    Lendet.top = 3360 - fsbVertical.Value
End Sub

Private Sub optMesme_Click()
    Call Pastro
    optSemestri1.Value = False
    optSemestri2.Value = False
    optPerfundimtare.Value = False
End Sub

Private Sub Pastro()
    txtAmzaNo.Text = ""
    txtEmri.Text = ""
    txtMbiemri.Text = ""
    txtAtesia.Text = ""
    txtKlasa.Text = ""
    txtIndeksi.Text = ""
    lblLargimi.Text = ""
    viti
    txtData.Visible = False
    Label3.Visible = False
    Lendet.Visible = False
    Cgrid1.Visible = False
    Label2.Visible = False
    lblNotatDheMungesat.Visible = False
    lblS1.Visible = False
    lblS2.Visible = False
    lblV.Visible = False
    lblR.Visible = False
    Label2.Visible = False
    lblNotatDheMungesat.Visible = False
    fsbVertical.Visible = False
    fsbHorizontali.Visible = False
End Sub

Private Sub optPerfundimtare_Click()
    Call pastrimi
End Sub

Private Sub optSemestri1_Click()
    Call pastrimi
End Sub

Private Sub optSemestri2_Click()
    Call pastrimi
End Sub

Private Sub optUlet_Click()
    Call Pastro
    optSemestri1.Value = False
    optSemestri2.Value = False
    optPerfundimtare.Value = False
End Sub

Private Sub pastrimi()
    txtData.Visible = False
    Label3.Visible = False
    Lendet.Visible = False
    Cgrid1.Visible = False
    Label2.Visible = False
    lblNotatDheMungesat.Visible = False
    lblS1.Visible = False
    lblS2.Visible = False
    lblV.Visible = False
    lblR.Visible = False
    fsbVertical.Visible = False
    fsbHorizontali.Visible = False
    cmdPrint.Enabled = False
End Sub

' Funksion qe merr si parameter nje numer dhe kthen ate numer ne
' shkronja qe i perkojne numrit ordinal te formatuar sipas defteses
Private Function MerrKlase() As String
    Dim numer As Integer
    On Error GoTo Gabimi
        numer = CInt(txtKlasa.Text)
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

Private Sub Text2_Change()

End Sub




