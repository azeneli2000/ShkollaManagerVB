VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMDIMain 
   BackColor       =   &H0080FF80&
   Caption         =   "Shkolla Manager"
   ClientHeight    =   8085
   ClientLeft      =   1995
   ClientTop       =   1455
   ClientWidth     =   8835
   Icon            =   "frmMDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMDIMain.frx":08CA
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   4320
      Top             =   5160
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbDown 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7710
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6244
            MinWidth        =   6244
            Text            =   "        Data e sotme"
            TextSave        =   "        Data e sotme"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlImages 
      Left            =   6360
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tblMenus 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   953
      ButtonWidth     =   661
      ButtonHeight    =   953
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "BT1"
            Key             =   "b1"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "s"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Bt2"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "s2"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuKonsultime 
      Caption         =   "Konsultime"
      Begin VB.Menu mnuKonfigurimeNotat 
         Caption         =   "Nxenesi"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuKonsultimeEvidenca 
         Caption         =   "Evidenca"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuKonsultimeAmza 
         Caption         =   "Amza"
         Shortcut        =   {F4}
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKonsDalje 
         Caption         =   "Dalje"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuHedhjeTeDhenash 
      Caption         =   "Hedhje Te Dhenash"
      Begin VB.Menu mnuHedhjeGjeneralitete 
         Caption         =   "Regjistrimi"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuHedhjeNotat 
         Caption         =   "Notat"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuHedhjeShenime 
         Caption         =   "Shenime"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuShenimeTePerkohshme 
         Caption         =   "Shenime te perkohshme"
      End
   End
   Begin VB.Menu mnuVeprime 
      Caption         =   "Veprime"
      Begin VB.Menu mnuVeprimeModifNota 
         Caption         =   "Modifiko Notat"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuVeprimeModGjeneralitete 
         Caption         =   "Modifiko Gjeneralitete"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuModifikoMungesa 
         Caption         =   "Modifiko Mungesa"
      End
      Begin VB.Menu modifikime_shenimesh 
         Caption         =   "Modifiko shenimet"
      End
      Begin VB.Menu mnuModifikoShenPerk 
         Caption         =   "Modifiko shënimet e përkohshme"
      End
      Begin VB.Menu mnuVeprimeEleminoNxenes 
         Caption         =   "Elemino Nxenes"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu mnuStatistika 
      Caption         =   "Statistika"
      Begin VB.Menu mnuStatistikaKlasat 
         Caption         =   "Klasat"
      End
      Begin VB.Menu mnuStatistikaNxenes 
         Caption         =   "Nxenesit"
      End
      Begin VB.Menu mnuMesataretMomentale 
         Caption         =   "Mesataret momentale"
         Begin VB.Menu mnuSipasKlasave 
            Caption         =   "Sipas klasave"
         End
         Begin VB.Menu mnuSipasCikleve 
            Caption         =   "Sipas cikleve"
         End
      End
      Begin VB.Menu mnuNxenesitDalluar 
         Caption         =   "Nxënësit e dalluar"
      End
   End
   Begin VB.Menu mnuKonfigurime 
      Caption         =   "Konfigurime"
      Begin VB.Menu mnuPerdorues 
         Caption         =   "Perdorues"
      End
      Begin VB.Menu mnuInformacione 
         Caption         =   "Informacione"
      End
   End
   Begin VB.Menu mnuInstrumente 
      Caption         =   "Instrumente"
      Begin VB.Menu mnuInstKlaseKonfig 
         Caption         =   "Konfigurim i lendeve"
      End
      Begin VB.Menu mnuInstKlaseKalo 
         Caption         =   "Kalo Klase"
      End
      Begin VB.Menu mnuFshi 
         Caption         =   "Fshi te dhenat"
      End
      Begin VB.Menu mnuKopjo 
         Caption         =   "Krijo back up"
      End
      Begin VB.Menu mnuKariko 
         Caption         =   "Kariko back up"
      End
      Begin VB.Menu mnuKonfigurimServeri 
         Caption         =   "Konfiguro Server"
      End
      Begin VB.Menu mnuTransferimi 
         Caption         =   "Sinkronizimi i te dhenave"
      End
      Begin VB.Menu DergoMungesat 
         Caption         =   "Dërgo mungesat"
      End
   End
   Begin VB.Menu mnuNdihme 
      Caption         =   "Ndihme "
      Begin VB.Menu mnuAktivizo 
         Caption         =   "Aktivizo programin"
      End
      Begin VB.Menu mnuNdihmePermbajtja 
         Caption         =   "Permbajtja"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuNdihmeRrethProgramit 
         Caption         =   "Rreth Programit"
      End
      Begin VB.Menu mnuKontakt 
         Caption         =   "Kontakt"
      End
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' * Author           :
' * Web Site         :
' * E-Mail           :
' * Date             : 11/08/2004
' * Time             : 15:24
' * Module Name      : frmMDIMain
' * Module Filenfame  : frmMDIMain.frm
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************


Option Explicit


Dim objUIController As clsUIController
Const SW_SHOWNORMAL = 1



Private Sub MDIForm_Load()
   loadIcons
   
   loadToolBar

   setIcon Me
   ' NESE PROGRAMI ESHTE VERSION I PLOTE ATEHERE C'AKTIVIZO MENUNE E AKTIVIZIMIT TE PROGRAMIT
   If lic.Stato = "LICENZA" Then
      Me.mnuAktivizo.Enabled = False
      Me.mnuFshi.Visible = False
      mnuFshi.Enabled = False
      Me.tblMenus.Buttons(2).Enabled = False
   End If
   resizeForm Me
   objectInitialization1 INICIALIZO_ADRESA_LOGO
   objectInitialization1 KONSULTIMI_I_DATESE_SE_MODIFIKIMIT
   percaktoTeDrejtat
   stbDown.Font.Bold = False
   stbDown.Panels(2).Text = "Data e modifikimit te fundit :    " & data_modifikimi
   stbDown.Panels(2).Alignment = sbrCenter
   
   stbDown.Panels(3).Text = "Data e sotme  :   " & date
   LexoServer
  ' viti
End Sub

Private Sub LexoServer()
Dim OggShell As WshShell
Set OggShell = New WshShell
'Dim MData As String
On Error Resume Next
Dim stringa As String
stringa = OggShell.RegRead("HKEY_CURRENT_USER\Software\" & "ShkollaManager" & "\Server")
If IsNull(stringa) Or stringa = "" Then
    stringa = InputBox("Jepni emrin (ose IP-ne) e serverit ku ndodhet baza e te dhenave te programit.", "Konfigurimi i serverit")
    If stringa = "" Then
        
    Else
        OggShell.RegWrite "HKEY_CURRENT_USER\Software\ShkollaManager\Server", stringa
        EmerServer = stringa
    End If
Else
    EmerServer = stringa
End If
End Sub

Private Sub loadToolBar()
   Dim objButton

   Set tblMenus.ImageList = imlImages

   With tblMenus.Buttons
      .Clear

      .Add(, "Exit", , , imlImages.ListImages("Exit").Index).ToolTipText = "Dalje"
      
        .Add(, "Aktivizo", , , imlImages.ListImages("aktivizo").Index).ToolTipText = "Aktivizo"
      .Add(, "Hedhje Te Dhenash - Nota", , , imlImages.ListImages("Insert").Index).ToolTipText = "Hedhje Te Dhenash - Nota"
     
      .Add(, "Hedhje Te Dhenash - Gjeneralitete", , , imlImages.ListImages("InsGjeneralitete").Index).ToolTipText = "Hedhje Te Dhenash - Gjeneralitete"

       .Add(, "Konsultime - Nota", , , imlImages.ListImages("regjister").Index).ToolTipText = "Konsultime - Nota"
     
     

      .Add(, "Konsultime - Evidenca", , , imlImages.ListImages("Evidenca").Index).ToolTipText = "Konsultime - Evidenca"
      .Add(, "Konsultime - Amza", , , imlImages.ListImages("modifikoGjeneralitete").Index).ToolTipText = "Konsultime - Amza"
      
      .Add(, "Modifikime - Nota", , , imlImages.ListImages("modifiko").Index).ToolTipText = "Modifikime - Nota"
      .Add(, "Modifikime - Gjeneralitete", , , imlImages.ListImages("modifikoGjeneralitete1").Index).ToolTipText = "Modifikime - Gjeneralitete"
      .Add(, "Elemino - Nxenes", , , imlImages.ListImages("elemNxenes").Index).ToolTipText = "Elemino - Nxenes"
      
      Set objButton = .Add(, "Statistika", , , imlImages.ListImages("statistika").Index)
      With objButton
          .ToolTipText = "Statistika"
          .ButtonMenus.Add 1, , "Klasa"
          .ButtonMenus.Add 2, , "Nxenes"
          .Style = tbrDropdown
      End With
      
      .Add(, "Hidh-Shenime", , , imlImages.ListImages("hidhshenime").Index).ToolTipText = "Hidh-Shenime"
      .Add(, "Modifiko-Shenime", , , imlImages.ListImages("modifikoshenime").Index).ToolTipText = "Modifiko-Shenime"
      .Add(, "Konfigurime - Perdorues", , , imlImages.ListImages("perdorues").Index).ToolTipText = "Konfigurime - Perdorues"
      .Add(, "Konfigurime - Informacione", , , imlImages.ListImages("perdorues1").Index).ToolTipText = "Konfigurime - Informacione"
      
      .Add(, "Konfiguro - Klase", , , imlImages.ListImages("konfigKlase").Index).ToolTipText = "Konfiguro - Klase"
      .Add(, "Kalo - Klase", , , imlImages.ListImages("kaloKlase").Index).ToolTipText = "Kalo - Klase"
      
      
      .Add(, "Help", , , imlImages.ListImages("help").Index).ToolTipText = "Help"
      .Add(, "Eksporto", , , imlImages.ListImages("Eksporto").Index).ToolTipText = "Eksporto ne Excel nxenesit e shkolles se mesme"
      .Add(, "EksportoExcel", , , imlImages.ListImages("EksportoExcel").Index).ToolTipText = "Eksporto raportin aktual në excel"
   End With

End Sub




Private Sub loadIcons()

   With imlImages.ListImages
      .Add 1, "Exit", LoadPicture(App.Path & "\Imazhe\exit1.bmp")
      .Add 2, "Insert", LoadPicture(App.Path & "\Imazhe\hedhje-nota.ico")
      .Add 3, "InsGjeneralitete", LoadPicture(App.Path & "\Imazhe\gjeneralitete.bmp")
      .Add 4, "regjister", LoadPicture(App.Path & "\Imazhe\amza-konsultime.ico")
      .Add 5, "modifiko", LoadPicture(App.Path & "\Imazhe\nota-modifikime.ico")
      .Add 6, "modifikoGjeneralitete", LoadPicture(App.Path & "\Imazhe\amza1.ico")
      .Add 7, "elemNxenes", LoadPicture(App.Path & "\Imazhe\delete.bmp")
      .Add 8, "statistika", LoadPicture(App.Path & "\Imazhe\statistik.ico")
      .Add 9, "perdorues", LoadPicture(App.Path & "\Imazhe\perdorues.ico")
      .Add 10, "perdorues1", LoadPicture(App.Path & "\Imazhe\inform.ico")
      .Add 11, "konfigKlase", LoadPicture(App.Path & "\Imazhe\konfig-lende.ico")
      .Add 12, "kaloKlase", LoadPicture(App.Path & "\Imazhe\PEOPLE.ICO")
      
      .Add 13, "help", LoadPicture(App.Path & "\Imazhe\help.ico")
      .Add 14, "Evidenca", LoadPicture(App.Path & "\\Imazhe\tileh.ico")
      .Add 15, "modifikoGjeneralitete1", LoadPicture(App.Path & "\\Imazhe\modifiko-gjeneralitete.ico")
      .Add 16, "hidhshenime", LoadPicture(App.Path & "\\Imazhe\hedhje-shenime.ico")
      .Add 17, "modifikoshenime", LoadPicture(App.Path & "\\Imazhe\modifikoshenime.ico")
      .Add 18, "aktivizo", LoadPicture(App.Path & "\\Imazhe\LOCK.ico")
      .Add 19, "Eksporto", LoadPicture(App.Path & "\\Imazhe\list.ico")
      .Add 20, "EksportoExcel", LoadPicture(App.Path & "\\Imazhe\Excel File (globe).ico")
   End With

End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
    Dim nResult As Integer
    If EmerServer = "" Then
        Cancel = 0
    End If
    nResult = MsgBox("Jeni te sigurte qe doni te dilni?", vbQuestion + vbYesNo, "Shkolla Manager")
    If nResult = 7 Then
        Cancel = 1
    Else: Cancel = 0
    End If
End Sub

Private Sub mnuAktivizo_Click()
   Dim Mp               As Processa
   Set Mp = New Processa
   Mp.CancellaRegistry App.Comments
   Mp.VerificaVersione 1022
End Sub

Private Sub mnuFshi_Click()
    Call fshiTeDhenat
End Sub

Private Sub mnuHedhjeGjeneralitete_Click()
   objectInitialization HEDHJE_TE_DHENASH_GJENERALITETE
End Sub

Private Sub mnuHedhjeNotat_Click()
  objectInitialization HEDHJE_TE_DHENASH_NOTA_MOMENTALEI
End Sub

Private Sub mnuHedhjeShenime_Click()
    objectInitialization HEDHJE_SJELLJE
End Sub

Private Sub mnuInformacione_Click()
  objectInitialization KONFIGURIME_INFORMACIONE
End Sub

Private Sub mnuInstKlaseKalo_Click()
   objectInitialization INSTRUMENTE_KLASE_KALO_KLASE
End Sub

Private Sub mnuInstKlaseKonfig_Click()
   objectInitialization INSTRUMENTE_KLASE_KONFIGURIM
End Sub

Private Sub mnuKariko_Click()
   
   Dim i                As Integer
   Dim gjendet          As Boolean
   i = MsgBox("Ju jeni duke zevendesuar te dhenat e programit me te dhenat e Back up-it." & Chr(10) & "Jeni te sigurte se doni te vazhdoni ?", vbQuestion + vbOKCancel, "Karikimi i te dhenave!")
   If i = vbOK Then
      gjendet = DirExists("C:\ShkollaManagerBackup")
      If gjendet Then
         If FileExists("C:\ShkollaManagerBackup\" & "DprogNec") Then
               Dim conn             As New ADODB.Connection
   
                With conn
                   .Provider = "Microsoft.Access.OLEDB.10.0"
                   .Properties("Data Provider").Value = "SQLOLEDB"
                   .Properties("Data Source").Value = EmerServer + "\SHM"
                   .Properties("User ID").Value = "sa"
                   .Properties("Password").Value = "vision"
                   
                   If (conn.state = 1) Then
                     .Close
                   End If
                   .Open
                End With
   
                Dim strSql As String
                'set offline
                strSql = "alter database DProgNec SET OFFLINE"
                conn.Execute strSql
                
                'restore
                strSql = "RESTORE DATABASE DProgNec FROM DISK = 'C:\ShkollaManagerBackup\" + "DProgNec" + "' WITH REPLACE"
                'conn.CommandTimeout = "120"
                conn.Execute strSql
                
                'set online
                strSql = "alter database DProgNec SET ONLINE"
                'conn.CommandTimeout = "30"
                conn.Execute strSql
                
                conn.Close
                
                MsgBox "Te dhenat e programit u kopjuan.", vbInformation, "Karikimi i back-up."
                
         Else
            MsgBox "Skedari  ' C:\ShkollaManagerBackup\DprogNec ' nuk ekziston .", vbInformation, "Karikimi i te dhenave !"
         End If
      Else
         MsgBox "Direktoria 'C:\ShkollaManagerBackup' ku duhet re ruhej skedari backup  nuk ekziston.", vbInformation, "Karikimi i te dhenave !"
      End If
   End If
End Sub
Private Sub MyTimerProc()

End Sub
Private Sub mnuKonfigurimeNotat_Click()
    objectInitialization KONSULTIME_NOTA_MOMENTALEI
End Sub

Private Sub mnuKonfigurimServeri_Click()
Dim OggShell As WshShell
Set OggShell = New WshShell
On Error Resume Next
Dim emerAktual As String
emerAktual = OggShell.RegRead("HKEY_CURRENT_USER\Software\" & "ShkollaManager" & "\Server")
Dim emerRi As String
emerRi = InputBox("Jepni emrin (ose IP-ne) e serverit ku ndodhet baza e të dhënave të programit Shkolla Manager.", "Shkolla Manager[Konfigurimi i serverit]", emerAktual)
    If Trim(emerRi) = "" Then
        OggShell.RegWrite "HKEY_CURRENT_USER\Software\ShkollaManager\Server", emerAktual
    Else
        
        On Error GoTo GabimLidhje:
        Dim conn             As New ADODB.Connection
        
        With conn
           .Provider = "Microsoft.Access.OLEDB.10.0"
           .Properties("Data Provider").Value = "SQLOLEDB"
           .Properties("Data Source").Value = emerRi + "\SHM"
           .Properties("User ID").Value = "sa"
           .Properties("Password").Value = "vision"
           .Properties("Initial Catalog").Value = "DProgNec"
           If (conn.state = 1) Then
             .Close
           End If
           .Open
        End With
                             
        OggShell.RegWrite "HKEY_CURRENT_USER\Software\ShkollaManager\Server", emerRi
        EmerServer = emerRi
        MsgBox "Emri i Serverit të Shkolla Manager u ruajt!", vbInformation
        Exit Sub
        
GabimLidhje:                 MsgBox "Nuk mund të kryhet lidhja me serverin e përcaktuar! Rishikoni emrin e serverit", vbCritical

    End If
End Sub

Private Sub mnuKonsDalje_Click()
   objectInitialization KONSULTIME_DALJE
End Sub

Private Sub mnuKonsultimeAmza_Click()
   objectInitialization KONSULTIME_AMZA
End Sub

Private Sub mnuKonsultimeEvidenca_Click()
   objectInitialization KONSULTIME_EVIDENCA
End Sub









Private Sub mnuKontakt_Click()
    Load frmSendMail
    frmSendMail.show vbModal
End Sub


Private Sub mnuKopjo_Click()

   Dim objKopjo         As New FileSystemObject
   If Not DirExists("C:\ShkollaManagerBackup") Then
      objKopjo.CreateFolder "C:\ShkollaManagerBackup"
   End If
   'objKopjo.CopyFile App.Path & "\ActiveX\" & "DprogNec.dat", "C:\ShkollaManagerBackup\" & "DprogNec.dat", True
   'MsgBox "Te dhenat e programit u kopjuan ne adresen :" & Chr(10) & "C:\ShkollaManagerBackup\" & "DprogNec.dat", vbInformation, "Ruani te dhenat."

   'Set objKopjo = Nothing
   Dim conn             As New ADODB.Connection
   
   With conn
      .Provider = "Microsoft.Access.OLEDB.10.0"
      .Properties("Data Provider").Value = "SQLOLEDB"
      .Properties("Data Source").Value = EmerServer + "\SHM"
      .Properties("User ID").Value = "sa"
      .Properties("Password").Value = "vision"
      .Properties("Initial Catalog").Value = "DProgNec"
      If (conn.state = 1) Then
        .Close
      End If
      .Open
   End With
   
   Dim strSql As String
   strSql = "BACKUP DATABASE DPrognec TO DISK = 'C:\ShkollaManagerBackup\" + "DProgNec" + "' WITH FORMAT"
   conn.Execute strSql
   MsgBox "Te dhenat e programit u kopjuan ne adresen :" & Chr(10) & "C:\ShkollaManagerBackup\" & "DprogNec", vbInformation, "Ruani te dhenat."
   conn.Close

End Sub



Private Sub mnuModifikoMungesa_Click()
    objectInitialization MODIFIKO_MUNGESA
End Sub

Private Sub mnuModifikoShenPerk_Click()
    objectInitialization MODIFIKO_SHENIME_PERKOHSHME
End Sub

Private Sub mnuNdihmePermbajtja_Click()
  CallHelp indeksHelp
End Sub

Private Sub mnuNdihmeRrethProgramit_Click()
   objectInitialization NDIHME_ABOUT
End Sub

Private Sub mnuNxenesitDalluar_Click()
    objectInitialization NXENESIT_DALLUAR
End Sub

Private Sub mnuPerdorues_Click()
   objectInitialization KONFIGURIME_PERDORUES
End Sub

Private Sub mnuShenimeTePerkohshme_Click()
    objectInitialization HEDHJA_SJELLJE_PERKOHSHME
End Sub

Private Sub mnuSipasCikleve_Click()
    objectInitialization MESATARET_CIKLI
End Sub

Private Sub mnuSipasKlasave_Click()
    objectInitialization MESATARET_MOMENTALE
End Sub

Private Sub mnuStatistikaKlasat_Click()
   objectInitialization STATISTIKA_KLASAT
End Sub

Private Sub mnuStatistikaNxenes_Click()
   objectInitialization STATISTIKA_NXENESIT
End Sub

Private Sub mnuTransferimi_Click()
    Dim hWnd
    Dim EkzProg
    If FileExists("C:\Program Files\TransferimiShkolla\Transferimi.exe") Then
        EkzProg = ShellExecute(hWnd, "open", "C:\Program Files\TransferimiShkolla\Transferimi.exe", "", "C:\", SW_SHOWNORMAL)
    ElseIf FileExists("C:\Programmi\TransferimiShkolla\Transferimi.exe") Then
        EkzProg = ShellExecute(hWnd, "open", "C:\Programmi\TransferimiShkolla\Transferimi.exe", "", "C:\", SW_SHOWNORMAL)
    Else
        CommonDialog1.DialogTitle = "Hapni skedarin Transferimi.exe"
        CommonDialog1.Filter = "Skedar i ekzekutueshem (Transferimi.exe)|Transferimi*.exe"
        CommonDialog1.ShowOpen
        If CommonDialog1.fileName <> "" Then
            EkzProg = ShellExecute(hWnd, "open", CommonDialog1.fileName, "", "C:\", SW_SHOWNORMAL)
        Else
            MsgBox "Transferimi i te dhenave nuk u krye!", vbExclamation, "Shkolla Manager"
        End If
        
    End If
End Sub

Private Sub DergoMungesat_Click()

    Dim nResult As Integer
    nResult = MsgBox("Jeni të sigurtë që doni të dërgoni mungesat me mesazh?", vbQuestion + vbYesNo, "Shkolla Manager")
    If (nResult = 6) Then
    Dim objFile As New FileSystemObject
        If (FileExists("C:\SkedaretTekst\tmp.txt")) Then
            objFile.DeleteFile ("C:\SkedaretTekst\tmp.txt")
        End If
        objFile.CreateTextFile ("C:\SkedaretTekst\tmp.txt")
        Dim Mp As Processa
        Set Mp = New Processa
        Dim sn As String
        sn = Mp.Leggi_Seriale
        Dim tmp As TextStream
        Set tmp = objFile.OpenTextFile("C:\SkedaretTekst\tmp.txt", ForWriting)
        tmp.WriteLine (sn)
        tmp.Close
        'ekzekuto programin per dergimin e mungesave me mesazh
        Dim hWnd
        Dim EkzProg
        'EkzProg = ShellExecute(hWnd, "open", "D:\DergoMungesaCs\DergoMungesaCs\bin\Debug\DergoMungesaCs.exe", "", "D:\", SW_SHOWNORMAL)
        'ndrysho
        If FileExists("C:\Program Files\DergoMungesat\DergoMungesaCs.exe") Then
            EkzProg = ShellExecute(hWnd, "open", "C:\Program Files\DergoMungesat\DergoMungesaCs.exe", "", "C:\", SW_SHOWNORMAL)
        ElseIf FileExists("C:\Programmi\DergoMungesat\DergoMungesaCs.exe") Then
            EkzProg = ShellExecute(hWnd, "open", "C:\Programmi\DergoMungesat\DergoMungesaCs.exe", "", "C:\", SW_SHOWNORMAL)
        Else
        MsgBox "Programi për dërgimin e mungesave me mesazh nuk është instaluar", vbExclamation, "Shkolla Manager"
        End If
    End If
End Sub


Private Sub mnuVeprimeEleminoNxenes_Click()
       objectInitialization VEPRIME_ELEMINO_NXENES
End Sub

Private Sub mnuVeprimeModGjeneralitete_Click()
   objectInitialization VEPRIME_MODIFIKO_GJENERALITETE
End Sub

Private Sub mnuVeprimeModifNota_Click()
   objectInitialization VEPRIME_MODIFIKO_NOTA
End Sub

Private Sub modifikime_shenimesh_Click()
    objectInitialization VEPRIME_MODIFIKO_SHENIME
End Sub


Private Sub tblMenus_ButtonClick(ByVal Button As MSComctlLib.Button)

   Select Case Button.Key
      Case "Exit"
         objectInitialization KONSULTIME_DALJE
      Case "Konsultime - Amza"
         objectInitialization KONSULTIME_AMZA
      Case "Konsultime - Evidenca"
         objectInitialization KONSULTIME_EVIDENCA
      Case "Hedhje Te Dhenash - Gjeneralitete"
         objectInitialization HEDHJE_TE_DHENASH_GJENERALITETE
      Case "Hedhje Te Dhenash - Nota"
         objectInitialization HEDHJE_TE_DHENASH_NOTA_MOMENTALEI
      Case "Modifikime - Nota"
         objectInitialization VEPRIME_MODIFIKO_NOTA
      Case "Modifikime - Gjeneralitete"
         objectInitialization VEPRIME_MODIFIKO_GJENERALITETE
      Case "Elemino - Nxenes"
         objectInitialization VEPRIME_ELEMINO_NXENES
      Case "Konfigurime - Perdorues"
         objectInitialization KONFIGURIME_PERDORUES
      Case "Konfigurime - Informacione"
         objectInitialization KONFIGURIME_INFORMACIONE
      Case "Konfiguro - Klase"
         objectInitialization INSTRUMENTE_KLASE_KONFIGURIM
      Case "Kalo - Klase"
         objectInitialization INSTRUMENTE_KLASE_KALO_KLASE
      Case "Konsultime - Nota"
         objectInitialization KONSULTIME_NOTA_MOMENTALEI
      Case "Hidh-Shenime"
         objectInitialization HEDHJE_SJELLJE
      Case "Modifiko-Shenime"
         objectInitialization VEPRIME_MODIFIKO_SHENIME
        
      Case "Hidh Nota Ne Amze"


      Case "Help"
         CallHelp 22

      Case "Eksporto"
         objectInitialization HEDHJA_NE_EKSEL
      Case "Aktivizo"
         Dim Mp               As Processa
         Set Mp = New Processa
         Mp.CancellaRegistry App.Comments
         Mp.VerificaVersione 1022
      Case "EksportoExcel"
         Call KonvertoExcel
   End Select

End Sub

Private Sub tblMenus_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
   Select Case ButtonMenu.Parent.Key
      'Case "Konsultime - Nota"
      '   Select Case ButtonMenu.Index
      '      Case 1:
      '         objectInitialization KONSULTIME_NOTA_MOMENTALEI
      '      Case 2:
      '         objectInitialization KONSULTIME_NOTA_SEMESTRI_I
      '      Case 3:
      '         objectInitialization KONSULTIME_NOTA_SEMESTRI_II
      '      Case 4:
      '         objectInitialization KONSULTIME_NOTA_VJETORE
      '   End Select
      Case "Statistika"
         Select Case ButtonMenu.Index
            Case 1:
               objectInitialization STATISTIKA_KLASAT
            Case 2:
               objectInitialization STATISTIKA_NXENESIT
         End Select
   End Select
End Sub


Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   If actionName = HEDHJA_NE_EKSEL Then
    objUIController.ExecuteActions
   End If
   objUIController.ExecuteActions True

   Set objUIController = Nothing
End Sub


Private Sub objectInitialization1(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
End Sub

Private Sub percaktoTeDrejtat()
    
    Select Case statusi
        
        Case "Vizitor"
            frmMDIMain.mnuHedhjeTeDhenash.Enabled = False
            frmMDIMain.mnuVeprime.Enabled = False
            'frmMDIMain.mnuKonfigurime.Enabled = False
            frmMDIMain.mnuInstrumente.Enabled = False
            frmMDIMain.mnuAktivizo.Enabled = False
            frmMDIMain.mnuInformacione.Enabled = False
            frmMDIMain.mnuFshi.Enabled = False
            frmMDIMain.mnuKopjo.Enabled = False
            frmMDIMain.mnuShenimeTePerkohshme.Enabled = False
            frmMDIMain.mnuKonfigurimServeri.Enabled = False
            frmMDIMain.mnuModifikoMungesa.Enabled = False
            mnuKariko.Enabled = False
            frmMDIMain.mnuPerdorues.Enabled = False
                        
            tblMenus.Buttons(2).Enabled = False
            tblMenus.Buttons(3).Enabled = False
            tblMenus.Buttons(4).Enabled = False
            tblMenus.Buttons(8).Enabled = False
            tblMenus.Buttons(9).Enabled = False
            tblMenus.Buttons(10).Enabled = False
            tblMenus.Buttons(12).Enabled = False
            tblMenus.Buttons(13).Enabled = False
            tblMenus.Buttons(14).Enabled = False
            tblMenus.Buttons(15).Enabled = False
            tblMenus.Buttons(16).Enabled = False
            tblMenus.Buttons(17).Enabled = False
            
            
            
            
        Case "SupervizorEmesme"
            
            frmMDIMain.mnuInformacione.Enabled = False
            frmMDIMain.mnuFshi.Enabled = False
            frmMDIMain.mnuKopjo.Enabled = False
            frmMDIMain.mnuKonfigurimServeri.Enabled = False
            mnuKariko.Enabled = False
        
        Case "SupervizorTetevjecare"
        
            frmMDIMain.mnuInformacione.Enabled = False
            frmMDIMain.mnuFshi.Enabled = False
            frmMDIMain.mnuKopjo.Enabled = False
            frmMDIMain.mnuKonfigurimServeri.Enabled = False
            mnuKariko.Enabled = False
            
        Case Else
    End Select
    
End Sub



  
