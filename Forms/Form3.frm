VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmZgjidhNxenes 
   Caption         =   "Zgjidh nxënës"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7305
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   7305
   StartUpPosition =   1  'CenterOwner
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridaNxenesit 
      Height          =   5055
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8916
      _Version        =   393216
      BackColorBkg    =   -2147483634
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H8000000E&
      DisabledPicture =   "Form3.frx":08CA
      DownPicture     =   "Form3.frx":6D0C
      Height          =   375
      Left            =   4320
      Picture         =   "Form3.frx":D14E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   2655
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H8000000E&
      DisabledPicture =   "Form3.frx":13590
      DownPicture     =   "Form3.frx":199D2
      Height          =   375
      Left            =   240
      Picture         =   "Form3.frx":1FE14
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   2655
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridaLabel 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   450
      _Version        =   393216
      BackColorBkg    =   -2147483634
      AllowBigSelection=   0   'False
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Ka disa nxënës me këto gjeneralitete. Zgjidhni njërin prej tyre!"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "frmZgjidhNxenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Amza As String
Public cikli As Integer

Private Sub cmdDil_Click()
Amza = ""
cikli = 0
Unload Me
End Sub

Private Sub cmdOK_Click()

If (Me.gridaNxenesit.row >= 0) Then
    Dim r As Integer
    r = Me.gridaNxenesit.row
    
    gridaNxenesit.col = 0
    Amza = Me.gridaNxenesit.Text
    
    gridaNxenesit.col = 2
    If (Me.gridaNxenesit.Text = "False") Then
        cikli = 0
    Else
        cikli = 1
    End If
    
    Unload Me
Else
    MsgBox "Zgjidhni njërin prej nxënësve në listë!", vbInformation, "Kujdes!"
End If
End Sub
