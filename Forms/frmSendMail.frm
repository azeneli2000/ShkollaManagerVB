VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSendMail 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dergo Mail"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   Icon            =   "frmSendMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   9135
   Begin MSWinsockLib.Winsock smtp 
      Left            =   480
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmSendMail.frx":0442
      DownPicture     =   "frmSendMail.frx":343F
      Height          =   375
      Left            =   3600
      Picture         =   "frmSendMail.frx":643C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      Caption         =   "Dil"
      DisabledPicture =   "frmSendMail.frx":9439
      DownPicture     =   "frmSendMail.frx":ECC3
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
      Left            =   6120
      Picture         =   "frmSendMail.frx":1454D
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog dialog 
      Left            =   5040
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "open file"
      Filter          =   "All files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmSendMail.frx":19DD7
      DownPicture     =   "frmSendMail.frx":1CE4E
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
      Left            =   600
      Picture         =   "frmSendMail.frx":1FEC5
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdBrowse 
      BackColor       =   &H80000009&
      Caption         =   "..."
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtAttachment 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1860
      TabIndex        =   10
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtMessage 
      Height          =   3975
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1920
      Width           =   7935
   End
   Begin VB.TextBox txtHost 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5880
      TabIndex        =   3
      Text            =   "80.78.66.66"
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtSender 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1860
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox txtRecipient 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1860
      TabIndex        =   1
      Text            =   "info@visioninfosolution.com"
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox txtSubject 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1860
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Bashkangjitje:"
      Height          =   195
      Left            =   840
      TabIndex        =   9
      Top             =   1320
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "SMTP Host:"
      Height          =   195
      Left            =   6960
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adresa juaj e e-mailit:"
      Height          =   195
      Left            =   330
      TabIndex        =   6
      Top             =   240
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marresi :"
      Height          =   195
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Subjekti:"
      Height          =   195
      Left            =   1185
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "frmSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
    dialog.ShowOpen
    txtAttachment.Text = dialog.fileName
End Sub

Private Sub cmdClear_Click()
    'Clear all fields
    Me.txtAttachment = ""
    Me.txtHost = "80.78.66.66"
    Me.txtMessage = ""
    Me.txtRecipient = "info@visioninfosolution.com"
    Me.txtSender = ""
    Me.txtSubject = ""
End Sub

Private Sub cmdDil_Click()
    Unload Me
End Sub

Private Sub cmdSend_Click()
    'Connect to the smtp server.
    'Smtp server port is everytime 25
    On Error GoTo fund
    If Not txtAttachment.Text = "" Then
        'EncodedFile = UUEncodeFile(txtAttachment)
    End If
    smtp.Connect txtHost.Text, 25
    'reset the state
    smtpState = MAIL_CONNECT
fund:
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ' Behet pozicionimi i formes ne mes te ekranit
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    ' Vendoset rendi i tabit
    txtSender.TabIndex = 1
    txtRecipient.TabIndex = 2
    txtSubject.TabIndex = 3
    txtAttachment.TabIndex = 4
    cmdSend.TabIndex = 5
    cmdClear.TabIndex = 6
    cmdDil.TabIndex = 7
End Sub

Private Sub smtp_DataArrival(ByVal bytesTotal As Long)
    Dim strServerResponse   As String
    Dim strResponseCode     As String
    Dim strDataToSend       As String
    'Retrive data from winsock buffer
    smtp.GetData strServerResponse
    Debug.Print strServerResponse
    'Get server response code (first three symbols)
    strResponseCode = Left(strServerResponse, 3)
    'Only these three codes from the server tell us
    'that the command was accepted
    If strResponseCode = "250" Or _
       strResponseCode = "220" Or _
       strResponseCode = "354" Then
        Select Case smtpState
            Case MAIL_CONNECT
                smtpState = MAIL_HELO
                'Remove blank spaces
                strDataToSend = Trim$(txtSender.Text)
                'Retrieve mailbox name from e-mail address
                strDataToSend = Left$(strDataToSend, _
                InStr(1, strDataToSend, "@") - 1)
                'Send HELO command to the server
                smtp.SendData "HELO " & strDataToSend & vbCrLf
                Debug.Print "HELO " & strDataToSend
            Case MAIL_HELO
                smtpState = MAIL_FROM
                'Send MAIL FROM command to the server
                'so it knows from who the message comes
                smtp.SendData "MAIL FROM:" & Trim$(txtSender.Text) & vbCrLf
                Debug.Print "MAIL FROM:" & Trim$(txtSender.Text)
            Case MAIL_FROM
                smtpState = MAIL_RCPTTO
                'Send RCPT TO command to the server
                'so it knows where to send the message
                smtp.SendData "RCPT TO:" & Trim$(txtRecipient.Text) & vbCrLf
                Debug.Print "RCPT TO:" & Trim$(txtRecipient.Text)
            Case MAIL_RCPTTO
                smtpState = MAIL_DATA
                'Send DATA command to the server
                'so it knows that we want to send the message
                smtp.SendData "DATA" & vbCrLf
                Debug.Print "DATA"
            Case MAIL_DATA
                smtpState = MAIL_DOT
                'Send Subject
                smtp.SendData "Subject:" & txtSubject.Text & vbLf & vbCrLf
                Debug.Print "Subject:" & txtSubject.Text
                Dim varLines    As Variant
                Dim varLine     As Variant
                Dim strMessage  As String
                'Add atacchments
                strMessage = txtMessage.Text & vbCrLf & vbCrLf & EncodedFile
                'clear the buffer for the encoded files
                EncodedFiles = ""
                'Parse message to get lines
                varLines = Split(strMessage, vbCrLf)
                'clear message buffer
                strMessage = ""
                'Send each line of the message
                'so no line gets lost
                For Each varLine In varLines
                    smtp.SendData CStr(varLine) & vbLf
                Next
                'Send a dot symbol so the server knows
                'that the end of the message is reached
                smtp.SendData "." & vbCrLf
                Debug.Print "."
            Case MAIL_DOT
                smtpState = MAIL_QUIT
                'Send QUIT command
                smtp.SendData "QUIT" & vbCrLf
                Debug.Print "QUIT"
            Case MAIL_QUIT
                'Close the connection to the smtp server
                smtp.Close
        End Select
    Else
        'Check if an error occured
        smtp.Close
        If Not smtpState = MAIL_QUIT Then
            'If yes then print the error
            MsgBox "Gabim: " & strServerResponse, vbCritical, "Gabim"
            Unload Me
        Else
            'if the message sent successfully, print it
            Unload Me
            Debug.Print "Mesazhi u dergua"
        End If
    End If
End Sub

Private Sub smtp_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Tell the user that an error occured
    If Description = "Address is not available from the local machine" Then
        Description = "Adresa juaj nuk eshte e sakte"
    End If
    MsgBox "Gabim numer " & Number & vbCrLf & Description, vbExclamation
End Sub

