Attribute VB_Name = "modFormLibrary"
' #VBIDEUtils#************************************************************
' * Author           :
' * Web Site         :
' * E-Mail           :
' * Date             : 11/08/2004
' * Time             : 15:41
' * Module Name      : modFormLibrary
' * Module Filename  : modFormLibrary.bas
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************
Public Type VektorNotash
    emer As String
    lende As String
    note As String
    date As Date
    llojNote As String
    semester As String
    Klasa As Integer
    indeksi As String
End Type

Public vektor() As VektorNotash



Public Function chkToBool(chkBox As CheckBox) As Boolean
   If (chkBox.Value = vbChecked) Then
      chkToBool = True
   Else
      chkToBool = False
   End If
End Function

Public Function BoolToBit(b As Boolean) As Integer
    If b Then
        BoolToBit = vbChecked
    Else
        BoolToBit = vbUnchecked
    End If
End Function

Public Function BoolToCheck(b As Boolean) As CheckBoxConstants
    If b Then
        BoolToCheck = vbChecked
    Else
        BoolToCheck = vbUnchecked
    End If
End Function

Public Sub setIcon(f As Form)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_setIcon

   Set f.Icon = LoadPicture(App.Path & "\Images\icon.ico")

EXIT_setIcon:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_setIcon:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in setIcon" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_setIcon
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_setIcon

End Sub


Public Function resizeForm(f As Form, Optional perqindje As Double = 1)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_resizeForm

   Dim C                As Control
   Dim koef             As Double
   koef = perqindje
   
   For Each C In f.Controls
      If TypeOf C Is Label Or TypeOf C Is TextBox Or TypeOf C Is Form Or TypeOf C Is Button Then 'Or TypeOf C Is SGGrid Then
         C.Width = C.Width * koef
         C.Height = C.Height * koef
      End If
   Next C
   
   f.Width = f.Width * koef
   f.Height = f.Height * koef

EXIT_resizeForm:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_resizeForm:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in resizeForm" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_resizeForm
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_resizeForm

End Function


Public Sub formPosition(objForm As Form)
  objForm.Top = frmMDIMain.tblMenus.Top
  'objForm.Left = frmMDIMain.iLstBar.Left
  objForm.Width = frmMDIMain.ScaleWidth
  objForm.Height = frmMDIMain.ScaleHeight
  'resizeForm objForm, 1.5
End Sub

Public Function selektoTextBox(txt As TextBox) As Boolean
   txt.SelStart = 0
   txt.SelLength = Len(txt)
   txt.SetFocus
End Function


Public Sub loadForm(objForm As Form)
   
   setIcon objForm
   
   formPosition objForm
   
   'objForm.Show
   
End Sub


Public Function indeksHelp() As Long
    Dim nr As Long
    Dim frmtmp As Form
    Dim s As String
    Set frmtmp = active_form
    
    If frmtmp Is Nothing Then
        indeksHelp = 1
        Exit Function
    End If
    s = frmtmp.Name
    Select Case s
        Case "frmHedhjeGjeneralitete"
            nr = 5
        Case "frmEleminoNxenes"
            nr = 3
        Case "Form1"
            nr = 6
        Case "frmShenime"
            nr = 24
        Case "frmInformacione"
            nr = 8
        Case "frmInstrumenteKaloKlase"
            nr = 10
        Case "frmInstrumenteLende"
            nr = 13
        Case "frmKonsultimeAmza"
            nr = 1
        Case "frmKonsultimeEvidenca"
            nr = 4
        Case "frmKonsultimeNota"
            nr = 20
        Case "frmModifikimeGjeneralitete"
            nr = 17
        Case "frmModifikimeNota"
            nr = 18
        Case "frmModifikimeSjellje"
            nr = 19
        Case "frmPerdorues"
            nr = 23
        Case "frmStatistikaKlasa"
            nr = 11
        Case "frmStatistikaNxenes"
            nr = 21
        Case Else
            nr = 16
    End Select
    indeksHelp = nr
    
End Function
