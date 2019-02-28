VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
' * Author           :
' * Web Site         :
' * E-Mail           :
' * Date             : 11/22/2004
' * Time             : 20:27
' * Module Name      : clsUserInterface
' * Module Filename  : clsUserInterface.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************



'local variable(s) to hold property value(s)
Private mvarActionUI As GUI_ACTION__ENUM
Private mvarErrorHandler As New clsErrorHandler



Public Property Set ErrorHandler(ByVal vData As clsErrorHandler)
    Set mvarErrorHandler = vData
End Property


Public Property Get ErrorHandler() As clsErrorHandler
   Set ErrorHandler = mvarErrorHandler
End Property


Public Property Let ActionUI(ByVal vData As GUI_ACTION__ENUM)
    mvarActionUI = vData
End Property


Public Property Get ActionUI() As GUI_ACTION__ENUM
   ActionUI = mvarActionUI
End Property


Private Sub Class_Terminate()
   Set ErrorHandler = Nothing
End Sub


Public Sub Execute_UIAction()
   Dim forma_e_vjeter As Form
   If Not active_form Is Nothing Then
      Set forma_e_vjeter = active_form
      'Unload active_form
      Set active_form = Nothing
   End If

   Select Case Me.ActionUI
      Case KONSULTIME_NOTA_MOMENTALEI

         Set active_form = New frmKonsultimeNota

      Case KONSULTIME_NOTA_MOMENTALEII

         Set active_form = New frmKonsultimeNota

      Case KONSULTIME_NOTA_MUNGESA

         Set active_form = New frmKonsultimeNota

      Case KONSULTIME_NOTA_SEMESTRI_I

         Set active_form = New frmKonsultimeNota

      Case KONSULTIME_NOTA_SEMESTRI_II

         Set active_form = New frmKonsultimeNota

      Case KONSULTIME_NOTA_VJETORE

         Set active_form = New frmKonsultimeNota

      Case KONSULTIME_EVIDENCA

         Set active_form = New frmKonsultimeEvidenca

      Case KONSULTIME_AMZA

         Set active_form = New frmKonsultimeAmza

      Case KONSULTIME_DALJE
        Dim nResult As Integer
            nResult = MsgBox("Jeni te sigurte qe doni te dilni?", vbQuestion + vbYesNo, "Shkolla Manager")
            If nResult = 7 Then
                
            Else: End
            End If

      Case HEDHJE_TE_DHENASH_GJENERALITETE

         Set active_form = New frmHedhjeGjeneralitete
         
      Case HEDHJE_TE_DHENASH_NOTA_MOMENTALEI
          
         Set active_form = New Form1
         
      Case VEPRIME_MODIFIKO_NOTA

         Set active_form = New frmModifikimeNota

      Case VEPRIME_MODIFIKO_GJENERALITETE

         Set active_form = New frmModifikimeGjeneralitete

      Case VEPRIME_ELEMINO_NXENES
         Set active_form = New frmEleminoNxenes
      Case STATISTIKA_KLASAT

         Set active_form = New frmStatistikaKlasa

      Case STATISTIKA_NXENESIT

         Set active_form = New frmStatistikaNxenes

      Case INSTRUMENTE_KLASE_KONFIGURIM

         Set active_form = New frmInstrumenteLende
      
      Case INSTRUMENTE_KLASE_KALO_KLASE
      
         Set active_form = New frmInstrumenteKaloKlase
         
      Case KONFIGURIME_PERDORUES
         
         Set active_form = New frmPerdorues
         
      Case KONFIGURIME_INFORMACIONE
      
         Set active_form = New frmInformacione
      
      Case NDIHME_ABOUT
      
         Set active_form = New frmAbout
      Case HEDHJE_SJELLJE
      
         Set active_form = New frmShenime
      Case HEDHJA_SJELLJE_PERKOHSHME
      
         Set active_form = New frmShenimeTePerkohshme
      Case VEPRIME_MODIFIKO_SHENIME
         
         Set active_form = New frmModifikimeSjellje
     Case MODIFIKO_SHENIME_PERKOHSHME
     
         Set active_form = New frmModifikoShenimeTePerkohshme
         
     Case MODIFIKO_MUNGESA
         Set active_form = New frmModifikoMungesa
     Case MESATARET_MOMENTALE
        Set active_form = New frmMesataretMomentale
     Case NXENESIT_DALLUAR
        Set active_form = New frmNxenesitDalluar
     Case MESATARET_CIKLI
        Set active_form = New frmMesataretCikli
     Case Backup
     
   End Select
   
   If Not active_form Is Nothing Then
      active_form.show
   End If
   
   If Not forma_e_vjeter Is Nothing Then
      Unload forma_e_vjeter
      
   End If
   
End Sub

