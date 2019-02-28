VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusManager"
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
' * Time             : 20:51
' * Module Name      : clsBusManager
' * Module Filename  : clsBusManager.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************

'local variable(s) to hold property value(s)
Private mvarErrorHandler As New clsErrorHandler
Private mvarActionName As GUI_ACTION__ENUM

Dim objDBManager As clsDBManager


Public Enum BUS_MANAGER_ENUM
   HIDH_PROVIME
   HIDH_NOTA
   GET_INFO_HEDHJE_NOTA 'lendet dhe nxenesit qe do hidhen notat
   GET_LENDE_AMZA
   RUAJ_AMZA
   RUAJ_PERDORUES_I_RI
   KERKO_PERDORUES
   MODIFIKO_PERDORUES
   CHECK_AMZA_REC
   RUAJ_AMZA_REC
   RUAJ_NOTA_REC
   RUAJ_NOTA_AMZA_REC
   RUAJ_NOTA_AMZA_REC1
   GET_AMZA_KEY
   GET_NOTA_KEY
   GET_NOTAAMZA_KEY
   MERR_NOTA_MOMENTALEI
   RUAJ_INFORMACIONE
   MERR_NOTA_AMZA
   MERR_NOTA_EVIDENCA
   NXENES_KALO_KLASE
   KALO_LENDE
   HIDH_GJENERALITETE
   MODIFIKO_NOTA
   MODIFIKO_AMZA1
   MODIFIKO_AMZA2
   LENDE_KLASA_KLIK
   LENDE_VITI_SHKOLLOR_KLIK
   KALO_KLASE_KERKO
   MERR_AMZA_E_MESME
   MERR_AMZA_TETEVJECARE
   SHFAQ_KLASAT_STATISTIKA
   SHFAQ_MESATARE_NXENESIT
   GABIME_KONSULTIME_NOTA
   GABIME_KONSULTIME_EVIDENCA
   SHFAQ_KLASAT_MESATARE
   SHFAQ_MESATARE_LENDET
   KALO_NE_LISTE_NXENES
   KERKO_NXENES_PER_LARGIM
   LARGO_NXENES
   SJELLJE_KERKO
   SJELLJE_HIDH
   LENDE_HIDH_PROVIME
   VERIFIKO_PERDORUES
   SHFAQ_SJELLJE
   HIDH_INFORMACIONE_NE_DATABASE
   HIDH_INFORMACIONE_NE_FORME
   SHENIME_LISTA_KLIK
   SHENIME_BUTONI_KLIK
   JEP_ADRESA_LOGO
   MODIFIKO_SHENIME_SJELLJE
   MODIFIKO_GJENERALITETE_SHENIME
   LOGIMI
   EKZISTON_PERDORUES
   ELIMINO_PERDORUES
   KERKO_KLASA_E_RE
   SHFAQ_KLASA_PROVIMET
   HIDH_PROVIME_NOTA
   PASTRO_KLASE
   HIDH_NOTA_MOMENTALE
   HIDH_NOTA_PERFUNDIMTARE
   HIDH_MUNGESA
   MODIFIKO_PROVIME
   NDRYSHO_PROVIME
   HIDH_NE_EKSEL
   VERIFIKO_NOTAT
   KONTROLLO_REGJISTRIMIN
   PRINTO_DEFTESA_SERI
   HIQ_LENDE
   SHFAQ_NOTA
   MERR_MUNGESA
   KONSULTO_DATEN
   MODIFIKO_DATEN
   FSHI_TE_DHENAT
   FSHI_PROVIMET
   NUMERO_NXENES
   SHENIM_PERKOHSHEM_KERKO
   SHENIME_PERKOHSHME_HIDH
   GJEJ_SHENIME_PERKOHSHME_PER_NXENESIN
   MODIFIKO_SHENIME_PERKOHSHME_PER_NXENESIN
   VIZUALIZO_MUNGESA_ACTION
   MUNGESA_ME_ARSYE_ACTION
   MUNGESA_PA_ARSYE_ACTION
   FSHI_MUNGESA_ACTION
   VIZUALIZO_MESATARET_MOMENTALE_ACTION
   VIZUALIZO_NXENESIT_DALLUAR_ACTION
   VIZUALIZO_MESATARET_CIKLI_ACTION
   SHFAQ_PROVIME_ACTION
   MODIFIKO_LENDE_ACTION
   HIDH_EKSEL_ACTION
End Enum



Public Property Let actionName(ByVal vData As GUI_ACTION__ENUM)
    mvarActionName = vData
End Property


Public Property Get actionName() As GUI_ACTION__ENUM
   actionName = mvarActionName
End Property


Public Property Set ErrorHandler(ByVal vData As clsErrorHandler)
    Set mvarErrorHandler = vData
End Property


Public Property Get ErrorHandler() As clsErrorHandler
   Set ErrorHandler = mvarErrorHandler
End Property


Private Sub createDBManager(paramActionName As BUS_MANAGER_ENUM)
   Set objDBManager = New clsDBManager
    objDBManager.actionName = paramActionName
End Sub


Private Sub destroyDBManager()
   Set objDBManager = Nothing
End Sub


Public Sub ExecuteActions()
   Select Case Me.actionName
     
         
       Case KONFIGURIME_PERDORUES_I_RI
         createDBManager RUAJ_PERDORUES_I_RI
         objDBManager.andiNew
        
       Case KONFIGURIME_PERDORUES_KERKO
         createDBManager KERKO_PERDORUES
         objDBManager.andiNew
        
       Case KONFIGURIME_PERDORUES_MODIFIKO
         createDBManager MODIFIKO_PERDORUES
         objDBManager.andiNew
        
       Case HEDHJE_NOTASH_INFO_KLASA
         createDBManager GET_INFO_HEDHJE_NOTA
         objDBManager.andiNew
        
       Case KONSULTIME_NOTA_MOMENTALEI
         createDBManager MERR_NOTA_MOMENTALEI
         objDBManager.andi
       Case KONFIGURIME_INFORMACIONE
         createDBManager RUAJ_INFORMACIONE
         objDBManager.andi
       Case KONSULTIME_AMZA
         createDBManager MERR_NOTA_AMZA
         objDBManager.andi
       Case VISUALIZO_EVIDENCE
         createDBManager MERR_NOTA_EVIDENCA
         objDBManager.andi
       Case INSTRUMENTE_KALO_KLASE
         createDBManager NXENES_KALO_KLASE
         objDBManager.andi
       Case INSTRUMENTE_LENDE
         createDBManager KALO_LENDE
         objDBManager.andi
       Case HEDHJE_GJENERALITETE
         createDBManager HIDH_GJENERALITETE
         objDBManager.andi
       Case MODIFIKO_NOTAT
         createDBManager MODIFIKO_NOTA
         objDBManager.andi
       Case MODIFIKO_GJENERALITETE1
         createDBManager MODIFIKO_AMZA1
         objDBManager.andiTjeter
       Case MODIFIKO_GJENERALITETE2
         createDBManager MODIFIKO_AMZA2
         objDBManager.andi
       Case INSTRUMENTE_LENDE_KLASA_KLIK
         createDBManager LENDE_KLASA_KLIK
         objDBManager.andi
       Case INSTRUMENTE_VITI_SHKOLLOR_KLIK
         createDBManager LENDE_VITI_SHKOLLOR_KLIK
         objDBManager.andi
       Case INSTRUMENTE_KALO_KLASE_KERKO
         createDBManager KALO_KLASE_KERKO
         objDBManager.andi
       Case KONSULTIME_AMZA_E_MESME
         createDBManager MERR_AMZA_E_MESME
         objDBManager.andi
       Case KONSULTIME_AMZA_TETEVJECARE
         createDBManager MERR_AMZA_TETEVJECARE
         objDBManager.andi
       Case STATISTIKA_NXENES_KLASAT
         createDBManager SHFAQ_KLASAT_STATISTIKA
         objDBManager.andi
       Case STATISTIKA_NXENES_MESATARE
         createDBManager SHFAQ_MESATARE_NXENESIT
         objDBManager.andi
         
        Case HEDHJE_TE_DHENASH_NOTA
          createDBManager HIDH_NOTA
          objDBManager.andi
        Case HEDHJE_TE_DHENASH_PROVIME
          createDBManager HIDH_PROVIME
          objDBManager.andi
        
        Case STATISTIKA_KLASA_SHFAQ
            createDBManager SHFAQ_KLASAT_MESATARE
            objDBManager.andi
        Case STATISTIKA_LENDET_MESATARET
            createDBManager SHFAQ_MESATARE_LENDET
            objDBManager.andi
        Case INSTRUMENTE_NXENES_KALO_NE_LISTE
            createDBManager KALO_NE_LISTE_NXENES
            objDBManager.andiTjeter
        Case ELIMINO_NXENES_KERKO
            createDBManager KERKO_NXENES_PER_LARGIM
            objDBManager.andiTjeter
        Case ELIMINO_NXENES_LARGO
            createDBManager LARGO_NXENES
            objDBManager.andiTjeter
        Case HEDHJE_SJELLJE_KERKO
            createDBManager SJELLJE_KERKO
            objDBManager.andiTjeter
        Case HEDHJA_SJELLJA_OK
            createDBManager SJELLJE_HIDH
            objDBManager.andiTjeter
        Case INSTRUMENTE_LENDE_HIDH_PROVIME
            createDBManager LENDE_HIDH_PROVIME
            objDBManager.andi
        Case VERIFIKIMI_I_PERDORUESIT
            createDBManager VERIFIKO_PERDORUES
            objDBManager.andiNew
        Case AMZA_SJELLJE
            createDBManager SHFAQ_SJELLJE
            objDBManager.andiTjeter
        Case INFORMACIONE_MBI_SHKOLLEN
            createDBManager HIDH_INFORMACIONE_NE_DATABASE
            objDBManager.andiNew
        Case SHFAQ_INFORMACION_MBI_SHKOLLEN
            createDBManager HIDH_INFORMACIONE_NE_FORME
            objDBManager.andiNew
        Case MODIFIKO_SHENIME_LISTA_KLIK
            createDBManager SHENIME_LISTA_KLIK
            objDBManager.andiNew
        Case MODIFIKO_SHENIME_BUTONI_KLIK
            createDBManager SHENIME_BUTONI_KLIK
            objDBManager.andiNew
        Case INICIALIZO_ADRESA_LOGO
            createDBManager JEP_ADRESA_LOGO
            objDBManager.andiNew
        Case MODIFIKO_SHENIME
            createDBManager MODIFIKO_SHENIME_SJELLJE
            objDBManager.andiNew
        Case MODIFIKO_SHENIME_SHFAQ
            createDBManager MODIFIKO_GJENERALITETE_SHENIME
            objDBManager.andiNew
        Case HYRJE_NE_PROGRAM
            createDBManager LOGIMI
            objDBManager.andiNew
        Case KA_PERDORUES
            createDBManager EKZISTON_PERDORUES
            objDBManager.andiNew
        Case ELIMINIMI_I_PERDORUESIT
            createDBManager ELIMINO_PERDORUES
            objDBManager.andiNew
        Case INSTRUMENTE_KERKO
            createDBManager KERKO_KLASA_E_RE
            objDBManager.andiNew
        Case VISUALIZO_KLASA_PROVIMET
            createDBManager SHFAQ_KLASA_PROVIMET
            objDBManager.andiNew
        Case HEDHJA_PROVIME
            createDBManager HIDH_PROVIME_NOTA
            objDBManager.andiNew
        Case INSTRUMENTE_PASTRO_KLASE
            createDBManager PASTRO_KLASE
            objDBManager.andiNew
        Case HEDHJA_NOTAVE_MOMENTALE
            createDBManager HIDH_NOTA_MOMENTALE
            objDBManager.andiNew
        Case HEDHJA_NOTAVE_PERFUNDIMTARE
            createDBManager HIDH_NOTA_PERFUNDIMTARE
            objDBManager.andiNew
        Case HEDHJA_E_MUNGESAVE
            createDBManager HIDH_MUNGESA
            objDBManager.andiNew
        
        Case MODIFIKIMI_I_PROVIMEVE
            createDBManager MODIFIKO_PROVIME
            objDBManager.andiNew
        
        Case NDRYSHIMI_I_PROVIMEVE
            createDBManager NDRYSHO_PROVIME
            objDBManager.andiNew
        
        Case HEDHJA_NE_EKSEL
            createDBManager HIDH_NE_EKSEL
            objDBManager.andiNew
        Case KONTROLLI_I_NOTAVE
            createDBManager VERIFIKO_NOTAT
            objDBManager.andiTjeter
            
        Case KONTROLLI_I_REGJISTRIMIT
            createDBManager KONTROLLO_REGJISTRIMIN
            objDBManager.andiNew
        
        Case PRINTO_EVIDENCA
            createDBManager PRINTO_DEFTESA_SERI
            objDBManager.andiNew
        Case HEQJA_E_LENDEVE
            createDBManager HIQ_LENDE
            objDBManager.andiNew
        Case SHFAQJA_E_NOTAVE
            createDBManager SHFAQ_NOTA
            objDBManager.andiTjeter
        Case MUNGESAT_ME_PA
            createDBManager MERR_MUNGESA
            objDBManager.andiNew
        Case KONSULTIMI_I_DATESE_SE_MODIFIKIMIT
            createDBManager KONSULTO_DATEN
            objDBManager.andiTjeter
        Case MODIFIKIMI_I_DATES
            createDBManager MODIFIKO_DATEN
            objDBManager.andiTjeter
        Case FSHIRJA_E_TE_DHENAVE
            createDBManager FSHI_TE_DHENAT
            objDBManager.andiTjeter
        Case FSHIRJA_E_PROVIMEVE
            createDBManager FSHI_PROVIMET
            objDBManager.andiTjeter
        Case NUMER_NXENESISH
            createDBManager NUMERO_NXENES
            objDBManager.andiTjeter
        Case HEDHJE_SHENIM_PERKOHSHEM_KERKO
            createDBManager SHENIM_PERKOHSHEM_KERKO
            objDBManager.andiTjeter
        Case SHENIME_PERKOHSHME_HIDH_OK
            createDBManager SHENIME_PERKOHSHME_HIDH
            objDBManager.andiTjeter
        Case MODIFIKO_SHENIME_PERKOHSHME_LISTA_KLIK
            createDBManager GJEJ_SHENIME_PERKOHSHME_PER_NXENESIN
            objDBManager.andiTjeter
        Case MODIFIKO_SHENIME_PERKOHSHME_PER_NXENESIN_OK
            createDBManager MODIFIKO_SHENIME_PERKOHSHME_PER_NXENESIN
            objDBManager.andiTjeter
        Case VIZUALIZO_MUNGESA
            createDBManager VIZUALIZO_MUNGESA_ACTION
            objDBManager.andiTjeter
        Case MUNGESA_ME_ARSYE
            createDBManager MUNGESA_ME_ARSYE_ACTION
            objDBManager.andiTjeter
        Case MUNGESA_PA_ARSYE
            createDBManager MUNGESA_PA_ARSYE_ACTION
            objDBManager.andiTjeter
        Case FSHI_MUNGESA
            createDBManager FSHI_MUNGESA_ACTION
            objDBManager.andiTjeter
        Case VIZUALIZO_MESATARET_MOMENTALE
            createDBManager VIZUALIZO_MESATARET_MOMENTALE_ACTION
            objDBManager.andiTjeter
        Case VIZUALIZO_NXENESIT_DALLUAR
            createDBManager VIZUALIZO_NXENESIT_DALLUAR_ACTION
            objDBManager.andiTjeter
        Case VIZUALIZO_MESATARET_CIKLI
            createDBManager VIZUALIZO_MESATARET_CIKLI_ACTION
            objDBManager.andiTjeter
        Case SHFAQ_PROVIME
            createDBManager SHFAQ_PROVIME_ACTION
            objDBManager.andiTjeter
        Case MODIFIKO_LENDE
            createDBManager MODIFIKO_LENDE_ACTION
            objDBManager.andiTjeter
        Case HIDH_EKSEL
            createDBManager HIDH_EKSEL_ACTION
            objDBManager.andiTjeter
   End Select
End Sub


