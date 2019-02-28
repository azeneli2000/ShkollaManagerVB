VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Hyrja"
   ClientHeight    =   3510
   ClientLeft      =   4065
   ClientTop       =   4455
   ClientWidth     =   5805
   Icon            =   "frmLogin.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5805
   Begin VB.TextBox txtPassword 
      BackColor       =   &H80000003&
      ForeColor       =   &H80000004&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2145
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   990
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Default         =   -1  'True
      DisabledPicture =   "frmLogin.frx":030A
      DownPicture     =   "frmLogin.frx":674C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2175
      Picture         =   "frmLogin.frx":CB8E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1785
      Width           =   2220
   End
   Begin VB.TextBox txtUserName 
      BackColor       =   &H80000003&
      ForeColor       =   &H80000004&
      Height          =   345
      Left            =   2145
      TabIndex        =   1
      Top             =   480
      Width           =   2325
   End
   Begin VB.CommandButton cmdAnullo 
      BackColor       =   &H80000009&
      Cancel          =   -1  'True
      DisabledPicture =   "frmLogin.frx":12FD0
      DownPicture     =   "frmLogin.frx":19412
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2175
      Picture         =   "frmLogin.frx":1F854
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2505
      Width           =   2220
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Fjalekalimi:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1005
      Width           =   1440
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Perdoruesi:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   270
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   495
      Width           =   1440
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objUIController As clsUIController

Private Sub cmdAnullo_Click()
    Dim I As Integer
    I = MsgBox("Doni te dilni nga programi ? ", vbYesNo, "Logimi")
    If I = vbYes Then
        End
    End If
End Sub

Private Sub cmdOK_Click()
    objectInitialization HYRJE_NE_PROGRAM
End Sub

Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions
   
   Set objUIController = Nothing
End Sub

Private Sub Form_Load()
    txtUserName.TabIndex = 0
    txtPassword.TabIndex = 1
    cmdOK.TabIndex = 2
    cmdAnullo.TabIndex = 3
    'frmMDIMain.Enabled = False
    Call SocketsInitialize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SocketsCleanup
End Sub
