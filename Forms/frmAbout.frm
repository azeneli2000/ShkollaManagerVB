VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Rreth Programit"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6420
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   1305
      ScaleWidth      =   1785
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   4200
      Picture         =   "frmAbout.frx":52EA
      Top             =   3240
      Width           =   2130
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tel : 04 251972"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Website : www.visioninfosolution.com"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Adresa Elektronike : info@visioninfosolution.com"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "                  Te githa te drejtat e rezervuara"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   6015
   End
   Begin VB.Label lblShenime 
      BackColor       =   &H00FFFFC0&
      Caption         =   "  Copyright  ©  2005      Vision Info Solution "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   6375
   End
   Begin VB.Label lblVersioni 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Versioni: 1.0"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Programi ""Shkolla Manager"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  loadForm Me
  
  'lblVersionNo.Caption = App.Major & App.Minor
  Me.Width = 6540
  Me.Height = 4860
End Sub

