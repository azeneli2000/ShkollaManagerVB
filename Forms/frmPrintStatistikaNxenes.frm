VERSION 5.00
Begin VB.Form frmPrintStatistikaNxenes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printo "
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   Icon            =   "frmPrintStatistikaNxenes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDil 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3360
      Picture         =   "frmPrintStatistikaNxenes.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      Picture         =   "frmPrintStatistikaNxenes.frx":674C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.OptionButton optMesatarja 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Printo mesataren e klases"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
   Begin VB.OptionButton optRaporti 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Printo raportin meshkuj - femra"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Value           =   -1  'True
      Width           =   2655
   End
End
Attribute VB_Name = "frmPrintStatistikaNxenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public choice As Integer

Private Sub cmdOK_Click()
    If optRaporti.Value = True Then
        choice = 1
    ElseIf optMesatarja.Value = True Then
        choice = 2
    Else: choice = 0
    End If
    Me.Hide
End Sub

Private Sub Form_Load()
    ' Behet pozicionimi i formes ne mes te ekranit
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub cmdDil_Click()
    choice = 0
    Me.Hide
End Sub

