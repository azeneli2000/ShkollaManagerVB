VERSION 5.00
Begin VB.Form frmPrintEvidenca 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printo"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5970
   Icon            =   "frmPrintEvidenca.frx":0000
   LinkTopic       =   "Printimi"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optNotaKlasa 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Printo notat per te gjithe nxenesit"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   2280
      Width           =   2535
   End
   Begin VB.OptionButton optNxSelected 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Printo notat e nxenesit te perzgjedhur"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.OptionButton optDeftesat 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Printo deftesat e klases"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.OptionButton optEvidencat 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Printo evidencen e klases"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton cmdDil 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3600
      Picture         =   "frmPrintEvidenca.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      Picture         =   "frmPrintEvidenca.frx":674C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "frmPrintEvidenca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nota As Integer

Private Sub cmdDil_Click()
    nota = 0
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    'Set objKonNota = CreateObject(frmKonsultimeNota)
    If optEvidencat.Value = True Then
        nota = 1
    ElseIf optDeftesat.Value = True Then
        nota = 2
    ElseIf Me.optNxSelected.Value = True Then
        nota = 3
    ElseIf Me.optNotaKlasa.Value = True Then
        nota = 4
    End If
    Me.Hide
End Sub

Private Sub Form_Load()
    ' Behet pozicionimi i formes ne mes te ekranit
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub

