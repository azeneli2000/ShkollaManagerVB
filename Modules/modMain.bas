Attribute VB_Name = "modMain"

' * Author           :
' * Web Site         :
' * E-Mail           :
' * Date             : 11/08/2004
' * Time             : 14:52
' * Module Name      : modMain
' * Module Filename  : modMain.bas
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************

Option Explicit
Public Type Licenza
        LeggiSeriale As Double
        Versione As Double
        NomeApplicazione As String
        NumeroAttivazione As Double
        Guid As String
        NumeroLic As Double
        RegistryNLic As String
        DataAttivazione As String
        Stato As String
        Stop As Boolean
End Type

Global lic As Licenza
Global DEMO As Boolean     ' nqs procedura eshte demo
Global FREE As Boolean     ' nqs procedura eshte free
Public EmerServer As String
Public cikel As Boolean
Public nrAmzaModifiko As String
Public emerModifiko As String
Public mbiemerModifiko As String
Public atesiaModifiko As String
Public klasaModifiko As String
Public indeksiModifiko As String
Public nrVitiShkollorModifiko As String
Public nrAmza As String
Public active_form As Form
Public Connection_String As String
Public matricaNotat(100, 100) As String
Public modifikoNotat(100, 100) As Integer
Public modifikoNota As Integer
Public rreshtiModifiko As Integer
Public shtyllaModifiko As Integer
Public adresaLogo As String
Public emerShkolla As String
Public adresaShkolla As String
Public qytetiShkolla As String
Public rrethiShkolla As String
Public telefoniShkolla As String
Public website As String
Public email As String
Public nr_lenda_momentale As Integer
Public nr_nxenesit_momentale As Integer
Public statusi As String
Public tipiModifikoProvime As Boolean
Public data_modifikimi As String
Public Type VektorTeDhenash
    TeDhenatSpecifike(20) As String
    TabelNotave(20, 4) As String
End Type
Public numerNxenesish As Integer
Public regjistruarNeNjeKlase As Boolean
Dim objUIController As clsUIController
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1




Public Sub Main()
LexoServer
If (EmerServer = "") Then
    Exit Sub
End If
Dim Mp As Processa
Set Mp = New Processa
Mp.VerificaVersione 1022
If lic.Stato = "DEMO" Then
    objectInitialization NUMER_NXENESISH
    If numerNxenesish > 30 Then
        'MsgBox "Numri i nxenesve te hedhur ne program eshte me i madh nga numri i lejuar." & vbCrLf & "Per te hedhur nxenes te tjere, ju duhet te rregjistroni programin", vbCritical, "Shkolla Manager"
        Mp.NrNxenesIMadh
        Exit Sub
    End If
End If
Dim objSecurity As Object
Dim objNrSerial As Object
Dim objRregjistrimi As Object
If Initialize() = False Then
    Exit Sub
End If
If lic.RegistryNLic <> "OEH-00000000000" Then
 objectInitialization KA_PERDORUES
Else
objectInitialization KA_PERDORUES
End If
Set Mp = Nothing
'objectInitialization KA_PERDORUES


End Sub

Private Function Initialize() As Boolean
    On Error GoTo Gabim
    Dim objSecurity As Object
    Dim objNrSerial As Object
    Dim objRregjistrimi As Object
    Set objSecurity = CreateObject("Securit.RegistrationClass")
    Set objNrSerial = CreateObject("Securit.NumriSerial")
    Set objRregjistrimi = CreateObject("Securit.Rregjistrimi")
    If objSecurity.CelesVlefshem() = True Then
        Initialize = True
        Exit Function
    End If
    objNrSerial.ShowDialog
    If objNrSerial.rregjistronr = False Then
        Initialize = False
        Exit Function
    End If
    objRregjistrimi.ShowDialog
    If objRregjistrimi.regsakte = False Then
        Initialize = False
        Exit Function
    End If
    If objRregjistrimi.regsakte = True Then
        Initialize = True
        Exit Function
    End If
    Initialize = False
    Exit Function
Gabim:
    MsgBox ("Nje gabim ndodhi  ")
    Initialize = False
    Exit Function
End Function

Private Sub LexoServer()
Dim OggShell As WshShell
Set OggShell = New WshShell
'Dim MData As String
On Error Resume Next
Dim stringa As String
stringa = OggShell.RegRead("HKEY_CURRENT_USER\Software\" & "ShkollaManager" & "\Server")
If IsNull(stringa) Or stringa = "" Then
    stringa = InputBox("Jepni emrin (ose IP-ne) e serverit ku ndodhet baza e të dhënave të programit Shkolla Manager.", "Shkolla Manager[Konfigurimi i serverit]")
    If stringa = "" Then
        'Unload Me
        EmerServer = stringa
    Else
        OggShell.RegWrite "HKEY_CURRENT_USER\Software\ShkollaManager\Server", stringa
        EmerServer = stringa
    End If
Else
    EmerServer = stringa
End If
End Sub
Public Function gjatesiRekordseti(str As String, conn As ADODB.Connection) As Integer

   Dim nr As Integer
   nr = 0
   
   Dim recset              As New ADODB.Recordset
   recset.Open str, conn, adOpenDynamic, adLockOptimistic
   Do While Not recset.EOF
    nr = nr + 1
    recset.MoveNext
   Loop
   
   recset.Close
   
   gjatesiRekordseti = nr
End Function

Public Function ktheCiklin(Klasa As String, nrVitiShkollor As String) As Boolean
        
    Dim cikli As Boolean
    Dim vitFillimi As Integer
    If (nrVitiShkollor <> "") Then
        vitFillimi = CInt(Mid(nrVitiShkollor, 1, 4))
    Else
        vitFillimi = 1000
    End If
    Select Case Klasa
        Case "1"
            cikli = False
        Case "2"
            cikli = False
        Case "3"
            cikli = False
        Case "4"
            cikli = False
        Case "5"
            cikli = False
        Case "6"
            cikli = False
        Case "7"
            cikli = False
        Case "8"
            cikli = False
        Case "9"
            If (vitFillimi >= 2008) Then
                cikli = False
            Else
                cikli = True
            End If
        Case "10"
            cikli = True
        Case "11"
            cikli = True
        Case "12"
            cikli = True
        Case Else
            cikli = True
    End Select
    
    ktheCiklin = cikli
        
        

End Function

Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
End Sub
Public Sub GoToWeb(psWebAddress As String)
   Dim lSuccess         As Long
   lSuccess = ShellExecute(0&, vbNullString, _
   psWebAddress, vbNullString, "C:\", SW_SHOWNORMAL)
End Sub

Public Function OpenEmailProgram(ByVal EmailAddress As String) As Boolean
    Dim res As Long
    res = ShellExecute(0&, "open", "mailto:" & EmailAddress, vbNullString, _
        vbNullString, vbNormalFocus)
    OpenEmailProgram = (res > 32)
End Function

Public Sub SendEmail(sDest As String, Optional sSubject As String, _
            Optional sBody As String, Optional sCC As String, Optional sBCC As String)
    ShellExecute 0, vbNullString, "mailto:" & sDest & "?subject=" & sSubject & _
        "&body=" & sBody & "&CC=" & sCC & "&BCC=" & sBCC, 0&, 0&, 1
End Sub

' Funksion qe kontrollon nese nje file ekziston. Ne kete rast kthen true, ne te kundert
' kthen False
Public Function FileExists(fileName As String) As Boolean
    On Error Resume Next
    ' Sigurohu qe parametri i dhene nuk eshte nje direktori
    FileExists = (GetAttr(fileName) And vbDirectory) = 0
End Function

' Funksion qe kontrollon nese nje direktori ekziston
Public Function DirExists(DirName As String) As Boolean
    On Error Resume Next
    'Sigurohu qe parametri i dhene nuk perfaqeson nje skedar
    DirExists = GetAttr(DirName) And vbDirectory
End Function

' Funksion qe verifikon nese draivi i flopit eshte gati
Public Function IsDriveReady(sDrive As String) As Boolean
    Dim fso As New FileSystemObject
    IsDriveReady = fso.GetDrive(sDrive).IsReady
    Set fso = Nothing
End Function


Public Sub KopjoSkedar(sourcePath As String, destPath As String)
    Dim fso As New FileSystemObject
    If FileExists(destPath + "amza.xls") Then
        fso.CopyFile sourcePath, destPath, True
    Else
        fso.CopyFile sourcePath, destPath, False
    End If
    Set fso = Nothing
   
End Sub

Public Function fshiTeDhenat()
    Dim i As Integer
    i = MsgBox("Ju jeni duke fshire te gjitha te dhenat ekzistuese !" & Chr(10) & "Doni te vazhdoni ? ", vbOKCancel + vbInformation, "Fshirja e te dhenave !")
    If i = vbOK Then
        objectInitialization FSHIRJA_E_TE_DHENAVE
    End If
    
    
End Function

Public Function formatoEmer(fjala As String) As String
    Dim i As Integer
    Dim str1 As String, str2 As String
    i = Len(fjala)
    If i = 0 Then
        formatoEmer = ""
    Else
        str1 = Mid(fjala, 1, 1)
        str2 = Mid(fjala, 2)
        formatoEmer = UCase(str1) & LCase(str2)
    End If
        
End Function



