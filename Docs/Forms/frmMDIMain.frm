VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Projekti StudManager"
   ClientHeight    =   6390
   ClientLeft      =   1995
   ClientTop       =   1455
   ClientWidth     =   8835
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbDown 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6015
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "11/03/2005"
         EndProperty
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
         Caption         =   "Notat"
      End
      Begin VB.Menu mnuKonsultimeEvidenca 
         Caption         =   "Evidenca"
      End
      Begin VB.Menu mnuKonsultimeAmza 
         Caption         =   "Amza"
      End
      Begin VB.Menu mnuKonsultPrinto 
         Caption         =   "Printo"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKonsDalje 
         Caption         =   "Dalje"
      End
   End
   Begin VB.Menu mnuHedhjeTeDhenash 
      Caption         =   "Hedhje Te Dhenash"
      Begin VB.Menu mnuHedhjeGjeneralitete 
         Caption         =   "Gjeneralitete"
      End
      Begin VB.Menu mnuHedhjeNotat 
         Caption         =   "Notat"
      End
      Begin VB.Menu mnuHedhjeShenime 
         Caption         =   "Shenime"
      End
   End
   Begin VB.Menu mnuVeprime 
      Caption         =   "Veprime"
      Begin VB.Menu mnuVeprimeModifNota 
         Caption         =   "Modifiko Notat"
      End
      Begin VB.Menu mnuVeprimeModGjeneralitete 
         Caption         =   "Modifiko Gjeneralitete"
      End
      Begin VB.Menu modifikime_shenimesh 
         Caption         =   "Modifiko shenimet"
      End
      Begin VB.Menu mnuVeprimeEleminoNxenes 
         Caption         =   "Elemino Nxenes"
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
   End
   Begin VB.Menu mnuNdihme 
      Caption         =   "Ndihme "
      Begin VB.Menu mnuAktivizo 
         Caption         =   "Aktivizo programin"
      End
      Begin VB.Menu mnuNdihmePermbajtja 
         Caption         =   "Permbajtja"
      End
      Begin VB.Menu mnuNdihmeIndeksi 
         Caption         =   "Indeksi"
      End
      Begin VB.Menu mnuNdihmeRrethProgramit 
         Caption         =   "Rreth Programit"
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
' * Module Filename  : frmMDIMain.frm
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************

Option Explicit


Dim objUIController As clsUIController

Private Sub MDIForm_Load()
   loadIcons

   loadToolBar

   setIcon Me

   resizeForm Me
   objectInitialization1 INICIALIZO_ADRESA_LOGO
  ' viti
End Sub

Private Sub loadToolBar()
   Dim objButton        As Button

   Set tblMenus.ImageList = imlImages

   With tblMenus.Buttons
      .Clear

      .Add(, "Exit", , , imlImages.ListImages("Exit").Index).ToolTipText = "Dalje"

      .Add(, "Hedhje Te Dhenash - Nota", , , imlImages.ListImages("Insert").Index).ToolTipText = "Hedhje Te Dhenash - Nota"
     
      .Add(, "Hedhje Te Dhenash - Gjeneralitete", , , imlImages.ListImages("InsGjeneralitete").Index).ToolTipText = "Hedhje Te Dhenash - Gjeneralitete"

       .Add(, "Konsultime - Nota", , , imlImages.ListImages("regjister").Index).ToolTipText = "Konsultime - Nota"
     
     

      .Add(, "Konsultime - Evidenca", , , imlImages.ListImages("Evidenca").Index).ToolTipText = "Konsultime - Evidenca"
      .Add(, "Konsultime - Amza", , , imlImages.ListImages("modifikoGjeneralitete").Index).ToolTipText = "Konsultime - Amza"
      
      .Add(, "Modifikime - Nota", , , imlImages.ListImages("modifiko").Index).ToolTipText = "Modifikime - Nota"
      .Add(, "Modifikime - Gjeneralitete", , , imlImages.ListImages("modifikoGjeneralitete").Index).ToolTipText = "Modifikime - Gjeneralitete"
      .Add(, "Elemino - Nxenes", , , imlImages.ListImages("elemNxenes").Index).ToolTipText = "Elemino - Nxenes"
      
      Set objButton = .Add(, "Statistika", , , imlImages.ListImages("statistika").Index)
      With objButton
          .ToolTipText = "Statistika"
          .ButtonMenus.Add 1, , "Klasa"
          .ButtonMenus.Add 2, , "Nxenes"
          .Style = tbrDropdown
      End With
      
      .Add(, "Konfigurime - Perdorues", , , imlImages.ListImages("perdorues").Index).ToolTipText = "Konfigurime - Perdorues"
      .Add(, "Konfigurime - Informacione", , , imlImages.ListImages("perdorues").Index).ToolTipText = "Konfigurime - Informacione"
      
      .Add(, "Konfiguro - Klase", , , imlImages.ListImages("konfigKlase").Index).ToolTipText = "Konfiguro - Klase"
      .Add(, "Kalo - Klase", , , imlImages.ListImages("kaloKlase").Index).ToolTipText = "Kalo - Klase"
      .Add(, "Hidh Nota Ne Amze", , , imlImages.ListImages("hidhNotaVjetore").Index).ToolTipText = "Hidh Nota Ne Amze"
      
      .Add(, "Printo", , , imlImages.ListImages("printo").Index).ToolTipText = "Printo"
      .Add(, "Help", , , imlImages.ListImages("help").Index).ToolTipText = "Help"

   End With

End Sub




Private Sub loadIcons()

   With imlImages.ListImages
      .Add 1, "Exit", LoadPicture(App.Path & "\Images\exit1.bmp")
      .Add 2, "Insert", LoadPicture(App.Path & "\Images\insert1.bmp")
      .Add 3, "InsGjeneralitete", LoadPicture(App.Path & "\Images\gjeneralitete.bmp")
      .Add 4, "regjister", LoadPicture(App.Path & "\Images\regjister.bmp")
      .Add 5, "modifiko", LoadPicture(App.Path & "\Images\insert1.bmp")
      .Add 6, "modifikoGjeneralitete", LoadPicture(App.Path & "\Images\gjeneralitete.bmp")
      .Add 7, "elemNxenes", LoadPicture(App.Path & "\Images\delete.bmp")
      .Add 8, "statistika", LoadPicture(App.Path & "\Images\statistika.ico")
      .Add 9, "perdorues", LoadPicture(App.Path & "\Images\perdorues.ico")
      .Add 10, "perdorues1", LoadPicture(App.Path & "\Images\perdorues.ico")
      .Add 11, "konfigKlase", LoadPicture(App.Path & "\Images\insert1.bmp")
      .Add 12, "kaloKlase", LoadPicture(App.Path & "\Images\insert1.bmp")
      .Add 13, "hidhNotaVjetore", LoadPicture(App.Path & "\Images\regjister.bmp")
      .Add 14, "help", LoadPicture(App.Path & "\Images\help.ico")
      .Add 15, "printo", LoadPicture(App.Path & "\Images\printo.bmp")
      .Add 16, "Evidenca", LoadPicture(App.Path & "\Images\tileh.ico")
   End With

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

Private Sub mnuKonfigurimeNotat_Click()
objectInitialization KONSULTIME_NOTA_MOMENTALEI
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









Private Sub mnuNdihmePermbajtja_Click()
  CallHelp indeksHelp
End Sub

Private Sub mnuNdihmeRrethProgramit_Click()
   objectInitialization NDIHME_ABOUT
End Sub

Private Sub mnuPerdorues_Click()
   objectInitialization KONFIGURIME_PERDORUES
End Sub

Private Sub mnuStatistikaKlasat_Click()
   objectInitialization STATISTIKA_KLASAT
End Sub

Private Sub mnuStatistikaNxenes_Click()
   objectInitialization STATISTIKA_NXENESIT
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

      Case "Hidh Nota Ne Amze"

      Case "Printo"

      Case "Help"

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
   objUIController.ExecuteActions True

   Set objUIController = Nothing
End Sub


Private Sub objectInitialization1(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
End Sub
