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


Public Enum ACTION_CASE_ENUM
   KONSULTIME_NOTA_MOMENTALE = 1
   KONSULTIME_NOTA_SEMESTRI_I = 2
   KONSULTIME_NOTA_SEMESTRI_II = 3
   KONSULTIME_NOTA_VJETORE = 4
   KONSULTIME_EVIDENCA = 5
   KONSULTIME_AMZA = 6
   KONSULTIME_PRINTO = 7
   KONSULTIME_DALJE = 8
   HEDHJE_TE_DHENASH_NOTA_MOMENTALE = 9
   HEDHJE_TE_DHENASH_NOTA_SEMESTRI_I = 10
   HEDHJE_TE_DHENASH_NOTA_SEMESTRI_II = 11
   HEDHJE_TE_DHENASH_NOTA_VJETORE = 12
   HEDHJE_TE_DHENASH_GJENERALITETE = 13
   VEPRIME_MODIFIKO_NOTA = 14
   VEPRIME_MODIFIKO_GJENERALITETE = 15
   VEPRIME_ELEMINO_NXENES = 16
   STATISTIKA_KLASAT = 17
   STATISTIKA_NXENESIT = 18
   PERDORUES_MODIFIKIME = 19
   PERDORUES_SHTO = 20
   INSTRUMENTE_KLASE_KONFIGURIM = 21
   INSTRUMENTE_KLASE_KALO_KLASE = 22
   INSTRUMENTE_HIDH_NOTA_VJETORE_NE_AMZE = 23
   NDIHME_PERMBAJTJA = 24
   NDIHME_INDEKSI = 25
   NDIHME_ABOUT = 26
End Enum

