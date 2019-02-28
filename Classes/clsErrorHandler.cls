VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsErrorHandler"
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
' * Time             : 20:53
' * Module Name      : clsErrorHandler
' * Module Filename  : clsErrorHandler.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************




'local variable(s) to hold property value(s)
Public mvarMesazhi As String
Public mvarGabimi As Integer


Public Sub set_gabimi(vData As Integer)
    mvarGabimi = vData
End Sub







Public Sub kapGabimin()
    
    Select Case active_form.Name
        Case "frmKonsultimeNota"
            kapGabimin_frmKonsultimeNota
        Case "frmKonsultimeEvidenca"
            kap_gabimin_frmKonsultimeEvidenca
        Case "frmKonsultimeAmza"
            kap_gabimin_frmKonsultimeAmza
        Case "frmHedhjeGjeneralitete"
            kap_gabimin_frmHedhjeGjeneralitete
        Case "Form1"
            kap_gabimin_frmHedhjeNota
        Case "frmModifikimeNota"
            kap_gabimin_frmModifikimeNota
        Case "frmStatistikaKlasa"
            kap_gabimin_frmStatistikaKlasa
        Case "frmStatistikaNxenes"
            kap_gabimin_frmStatistikaNxenes
        Case "frmInstrumenteKaloKlase"
            kap_gabimin_frmInstrumenteKaloKlase
        Case "frmModifikimeGjeneralitete"
            kap_gabimin_frmModifikimeGjeneralitete
        Case "frmEleminoNxenes"
            kap_gabimin_frmEleminoNxenes
        Case "frmModifikoMungesa"
            kap_gabimin_frmModifikoMungesa
        Case "frmMesataretMomentale"
            kap_gabimin_frmMesataretMomentale
        Case "frmNxenesitDalluar"
            kap_gabimin_frmNxenesitDalluar
        Case "frmMesataretCikli"
            kap_gabimin_frmMesataretCikli
    End Select
End Sub

Private Sub kap_gabimin_frmMesataretCikli()
If Not (active_form.optSemestri1.Value Or active_form.optSemestri2.Value) Then
    mvarGabimi = 1
    Exit Sub
End If
mvarGabimi = 0
End Sub
Private Sub kap_gabimin_frmNxenesitDalluar()
    If Not (active_form.optMesme.Value Or active_form.optUlet.Value) Then
        mvarGabimi = 1
        Exit Sub
    End If
    
    If (active_form.cboVitiShkollor.Text = "") Then
        mvarGabimi = 3
        Exit Sub
    End If
    
    If (active_form.cbNota.Value = 0 And active_form.cbMesatarja.Value = 0) Then
        mvarGabimi = 4
        Exit Sub
    End If
    
    If (active_form.cbMesatarja.Value And active_form.txtMesatarja.Text = "") Then
        mvarGabimi = 5
        Exit Sub
    End If
    
    If (active_form.cbMesatarja.Value) Then
        'kontrollo formatin e dhene te mesatares
        Dim mesatarja As String
        mesatarja = active_form.txtMesatarja.Text
        If (Len(mesatarja) > 4 Or Len(mesatarja) < 3) Then
            mvarGabimi = 6
            Exit Sub
        End If
        
        If (Len(mesatarja) = 4 And mesatarja <> "10.0") Then
            mvarGabimi = 6
            Exit Sub
        End If
        
        If (Mid(mesatarja, 2, 1) <> "." And Len(mesatarja) = 3) Then
            mvarGabimi = 6
            Exit Sub
        End If
        
        If (CInt(Mid(mesatarja, 1, 1)) < 5 And Len(mesatarja) = 3) Then
            mvarGabimi = 7
            Exit Sub
        End If
    End If
    mvarGabimi = 0
End Sub
Private Sub kapGabimin_frmKonsultimeNota()


   If Not (active_form.optMesme.Value Or active_form.optUlet.Value) Then
      mvarGabimi = 1
      Exit Sub
   End If
   If Not (active_form.optSemestri1.Value Or active_form.optSemestri2.Value Or active_form.optPerfundimtare.Value) Then
     mvarGabimi = 151
     Exit Sub
   End If
   If active_form.txtAmzaNo.Text = "" And (active_form.txtEmri.Text = "" And active_form.txtMbiemri.Text = "" And active_form.txtAtesia.Text = "") Then
      mvarGabimi = 2
      Exit Sub
     
   End If
   If active_form.cboVitiShkollor.Text = "" Then
     mvarGabimi = 2
     Exit Sub
   End If

   mvarGabimi = 0

End Sub
Private Sub kap_gabimin_frmKonsultimeEvidenca()
    
    If Not (active_form.optSemestri1 Or active_form.optSemestri2 Or active_form.optVjetore) Then
        mvarGabimi = 3
        Exit Sub
    End If
    If active_form.cboVitiShkollor.Text = "" Or active_form.cboKlasa.Text = "" Or active_form.cboIndeksi.Text = "" Then
        mvarGabimi = 4
        Exit Sub
    End If
    mvarGabimi = 0
End Sub

Private Sub kap_gabimin_frmMesataretMomentale()
    If Not (active_form.optSemestri1 Or active_form.optSemestri2) Then
        mvarGabimi = 1
        Exit Sub
    End If
    
    If active_form.cboVitiShkollor.Text = "" Or active_form.cboKlasa.Text = "" Or active_form.cboIndeksi.Text = "" Then
        mvarGabimi = 2
        Exit Sub
    End If
    mvarGabimi = 0
End Sub

Private Sub kap_gabimin_frmModifikoMungesa()
    If active_form.cboVitiShkollor.Text = "" Or active_form.cboKlasa.Text = "" Or active_form.cboIndeksi.Text = "" Then
        mvarGabimi = 1
        Exit Sub
    End If
    If active_form.dtpNga.Value > active_form.dtpDeri.Value Then
        mvarGabimi = 2
        Exit Sub
    End If
    
    Dim vitiMin, vitiMax As Integer
    Dim dtMin, dtMax As Date
    vitiMin = CInt(Mid(active_form.cboVitiShkollor.Text, 1, 4))
    vitiMax = CInt(Mid(active_form.cboVitiShkollor.Text, 6, 4))
    
    dtMin = CDate("01/08/" + CStr(vitiMin))
    dtMax = CDate("31/07/" + CStr(vitiMax))
    If (active_form.dtpNga.Value < dtMin Or active_form.dtpNga.Value > dtMax Or active_form.dtpDeri.Value < dtMin Or active_form.dtpDeri.Value > dtMax) Then
        mvarGabimi = 3
        Exit Sub
    End If
    Dim dif As Integer
    dif = active_form.dtpDeri.Value - active_form.dtpNga.Value
    If (dif > 15) Then
        mvarGabimi = 4
        Exit Sub
    End If
    Select Case statusi
        Case "SupervizorEmesme"
            If (ktheCiklin(active_form.cboKlasa.Text, CStr(active_form.cboVitiShkollor.Text)) = False) Then
                mvarGabimi = 5
                Exit Sub
            End If
        Case "SupervizorTetevjecare"
            If (ktheCiklin(active_form.cboKlasa.Text, CStr(active_form.cboVitiShkollor.Text)) = True) Then
                mvarGabimi = 5
                Exit Sub
            End If
    End Select
    mvarGabimi = 0
End Sub

Private Sub kap_gabimin_frmKonsultimeAmza()
   If Not (active_form.optMesme.Value Or active_form.optUlet.Value) Then
      mvarGabimi = 5
      Exit Sub
   End If

   If active_form.txtAmzaNo.Text = "" And (active_form.txtEmri.Text = "" And active_form.txtMbiemri.Text = "" And active_form.txtAtesia.Text = "") Then
   
      mvarGabimi = 6
      Exit Sub
     
   End If
   mvarGabimi = 0
End Sub
Private Sub kap_gabimin_frmHedhjeGjeneralitete()
    If active_form.txtAmzaNo.Text = "" Or active_form.txtEmri.Text = "" Or active_form.txtMbiemri.Text = "" Or active_form.cboKlasa.Text = "" Or active_form.cboIndeksi.Text = "" Or active_form.txtAtesia.Text = "" Or active_form.txtMemesia.Text = "" Or active_form.cboSeksi.Text = "" Or active_form.txtVendlindja = "" Or active_form.cboVitiShkollor = "" Or active_form.txtData.Visible = True Then
        mvarGabimi = 7
        Exit Sub
    End If
    If kontrollo_numerAmze(active_form.txtAmzaNo.Text) Then
        mvarGabimi = 8
        Exit Sub
    End If
    mvarGabimi = 0
End Sub
Private Sub kap_gabimin_frmHedhjeNota()
   If Not (active_form.optMomentale.Value Or active_form.optSemestrale.Value Or active_form.optSemestri1.Value Or active_form.optSemestri2.Value Or active_form.optVjetore.Value Or active_form.optDetyreKontrolli.Value Or active_form.optMungeseMe.Value Or active_form.optMungesePa.Value Or active_form.optLirimi.Value Or active_form.optMature.Value Or active_form.optRiprovim.Value) Then
      mvarGabimi = 9
      Exit Sub
   End If
   If (active_form.optSemestri1.Value Or active_form.optSemestri2.Value) And (active_form.optMomentale.Value = False And active_form.optDetyreKontrolli.Value = False And active_form.optMungeseMe.Value = False And active_form.optMungesePa.Value = False And active_form.optSemestrale.Value = False) Then
      mvarGabimi = 113
      Exit Sub
   End If
   If active_form.cboKlasa.Text = "" Or active_form.cboIndeksi.Text = "" Or active_form.cboVitiShkollor.Text = "" Then
      mvarGabimi = 10
      Exit Sub
   End If
   Dim vitFillimi As Integer
   vitFillimi = CInt(Mid(active_form.cboVitiShkollor.Text, 1, 4))
   Dim klaseProvim As String
   If (vitFillimi <= 2006) Then
        klaseProvim = "8"
    Else
        klaseProvim = "9"
   End If
   If active_form.optLirimi = True And active_form.cboKlasa.Text <> klaseProvim Then
      mvarGabimi = 111
      Exit Sub
   End If
   If active_form.optMature = True And active_form.cboKlasa.Text <> "12" Then
      mvarGabimi = 112
      Exit Sub
   End If
   mvarGabimi = 0
End Sub
Private Sub kap_gabimin_frmModifikimeNota()
    If Not (active_form.optMesme.Value Or active_form.optUlet.Value) Then
      mvarGabimi = 11
      Exit Sub
   End If
   If Not (active_form.optNota.Value Or active_form.optProvimi.Value) Then
      mvarGabimi = 1001
      Exit Sub
   End If
   If active_form.txtAmzaNo.Text = "" And (active_form.txtEmri.Text = "" And active_form.txtMbiemri.Text = "" And active_form.txtAtesia.Text = "") Then
   
      mvarGabimi = 12
      Exit Sub
     
   End If
   If active_form.cboVitiShkollor.Text = "" Then
     mvarGabimi = 12
     Exit Sub
   End If

   mvarGabimi = 0
End Sub

Private Sub kap_gabimin_frmStatistikaKlasa()
    If active_form.cboOptions.Text = "" Or active_form.cboVitiShkollor.Text = "" Then
         mvarGabimi = 13
         Exit Sub
    End If
    
    mvarGabimi = 0
End Sub
Private Sub kap_gabimin_frmStatistikaNxenes()
   If active_form.cboOptions.Text = "" Or active_form.cboVitiShkollor.Text = "" Then
      mvarGabimi = 14
      Exit Sub
   End If

   mvarGabimi = 0
End Sub
Private Sub kap_gabimin_frmInstrumenteKaloKlase()
    If Not (active_form.optUlet.Value Or active_form.optMesme.Value) Then
        mvarGabimi = 15
        Exit Sub
    End If
    
    mvarGabimi = 0
End Sub
Private Sub kap_gabimin_frmModifikimeGjeneralitete()
    If Not (active_form.optMesme.Value Or active_form.optUlet.Value) Then
      mvarGabimi = 17
      Exit Sub
   End If

   If active_form.txtAmzaNo.Text = "" And (active_form.txtEmri.Text = "" And active_form.txtMbiemri.Text = "" And active_form.txtAtesia.Text = "") Then
   
      mvarGabimi = 18
      Exit Sub
     
   End If
   If active_form.txtAmzaNo.Text = "" Or active_form.txtEmri.Text = "" Or active_form.txtMbiemri.Text = "" Or active_form.cboKlasa.Text = "" Or active_form.cboIndeksi.Text = "" Or active_form.txtAtesia.Text = "" Or active_form.txtMemesia.Text = "" Or active_form.cboSeksi.Text = "" Or active_form.txtVendlindja = "" Then
        mvarGabimi = 19
        Exit Sub
    End If
    If kontrollo_numerAmze(active_form.txtAmzaNo.Text) Then
        mvarGabimi = 20
        Exit Sub
    End If
    If (cikel <> ktheCiklin(active_form.cboKlasa.Text, CStr(active_form.cboVitiShkollor.Text))) Then
        mvarGabimi = 21
        Exit Sub
    End If
    
    mvarGabimi = 0
End Sub
Private Sub kap_gabimin_frmEleminoNxenes()
    If Not (active_form.optMesme.Value Or active_form.optTetevjecare.Value) Then
        mvarGabimi = 21
        Exit Sub
    End If
    If active_form.txtAmzaNo.Text = "" And (active_form.txtEmri.Text = "" And active_form.txtMbiemri.Text = "" And active_form.txtAtesia.Text = "") Then
   
      mvarGabimi = 22
      Exit Sub
     
   End If
   mvarGabimi = 0
End Sub


Public Sub menazhim_gabimi()
    Select Case active_form.Name
          Case "frmKonsultimeNota"
              menazhim_gabimi_konsultime_nota
          Case "frmKonsultimeEvidenca"
              menazhim_gabimi_konsultime_evidenca
          Case "frmKonsultimeAmza"
              menazhim_gabimi_konsultime_amza
          Case "frmHedhjeGjeneralitete"
              menazhim_gabimi_frmHedhjeGjeneralitete
          Case "Form1"
              menazhim_gabimi_frmHedhjeNota
          Case "frmModifikimeNota"
              menazhim_gabimi_frmModifikimeNota
          Case "frmStatistikaKlasa"
              menazhim_gabimi_frmStatistikaKlasa
          Case "frmStatistikaNxenes"
              menazhim_gabimi_frmStatistikaNxenes
          Case "frmInstrumenteKaloKlase"
              menazhim_gabimi_frmInstrumenteKaloKlase
          Case "frmModifikimeGjeneralitete"
              menazhim_gabimi_frmModifikimeGjeneralitete
          Case "frmEleminoNxenes"
              menazhim_gabimi_frmEleminoNxenes
          Case "frmModifikoMungesa"
              menazhim_gabimi_frmModifikoMungesa
          Case "frmMesataretMomentale"
              menazhim_gabimi_frmMesataretMomentale
          Case "frmNxenesitDalluar"
              menazhim_gabimi_frmNxenesitDalluar
          Case "frmMesataretCikli"
              menazhim_gabimi_frmMesataretCikli
    End Select
End Sub
    
Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
End Sub

Private Sub Class_Initialize()
    mvarGabimi = 0
    mvarMesazhi = ""
End Sub

Private Sub menazhim_gabimi_konsultime_nota()

   Select Case mvarGabimi
      Case 1
         MsgBox " Ju duhet te zgjidhni nje nga opsionet e shkolles se mesme ose nëntëvjeçare ", vbInformation, "Konsultimi i notave."

      Case 2
         MsgBox "Para se te konsultoni notat jepni vitin shkollor" + Chr(10) + "si edhe te paktën një nga të dhenat identifikuese per nxenesin " & Chr(10) & "qe mund te jene numri i amzes,emri, mbiemri ose atesia!", vbInformation, "Konsultimi i notave."

      Case 151
         MsgBox "Ju duhet te zgjidhni opsionin qe percakton" & Chr(10) & "llojin e notave qe doni te konsultoni.", vbExclamation, "Konsultimi i notave."
         
        

      Case Else

   End Select

End Sub

Private Sub menazhim_gabimi_konsultime_evidenca()
     Select Case mvarGabimi
        Case 3
            MsgBox "Ju duhet te zgjidhni nje nga opsionet Semestri1 , Semestri2 ose Vjetore " & Chr(10) & "para se te konsultoni evidencen per nje klase", vbInformation, "Konsultimi i evideces."
        Case 4
            MsgBox "Ju duhet te jepni te dhenat identifikuese te nje klase " & Chr(10) & "qe jane numri i klases, indeksi dhe viti shkollor para se te konsultoni evidencen per ate klase.", vbInformation, "Konsultimi i evidences."
        Case Else
    End Select
End Sub

Private Sub menazhim_gabimi_frmMesataretMomentale()
    Select Case mvarGabimi
        Case 1
            MsgBox "Ju duhet te zgjidhni nje nga opsionet Semestri1 ose Semestri2 " & Chr(10) & "para se te konsultoni mesataret momentale per nje klase!", vbInformation, "Mesataret momentale sipas klasave."
        Case 2
            MsgBox "Ju duhet te jepni te dhenat identifikuese te nje klase " & Chr(10) & "qe jane numri i klases, indeksi dhe viti shkollor para se te konsultoni mesataret momentale per ate klase.", vbInformation, "Mesataret momentale sipas klasave."
        Case Else
    End Select
End Sub

Private Sub menazhim_gabimi_konsultime_amza()
    
    Select Case mvarGabimi
      Case 5
         MsgBox " Ju duhet te zgjidhni nje nga opsionet e shkolles se mesme ose nëntëvjeçare ", vbInformation, "Konsultimi i amzes."

      Case 6
         MsgBox "Para se te konsultoni amzën jepni vitin shkollor" + Chr(10) + "si edhe te paktën një nga të dhenat identifikuese per nxenesin " & Chr(10) & "qe mund të jenë numri i amzës,emri, mbiemri ose atësia!", vbInformation, "Konsultimi i notave."
      Case Else

   End Select

End Sub
Private Sub menazhim_gabimi_frmHedhjeGjeneralitete()
    Select Case mvarGabimi
        Case 7
            MsgBox " Ju duhet te jepni te gjitha te dhenat e kerkuara me lart para se te hidhni nje nxenes te ri ne amze.", vbInformation, "Hedhja e gjeneraliteteve."
        Case 8
            MsgBox "Numri i amzes i dhene nga ju nuk eshte ne formatin e duhur.Ai duhet te permbaje vetem shifra.", vbInformation, "Hedhja e gjeneraliteteve."
    End Select
End Sub
Private Sub menazhim_gabimi_frmHedhjeNota()
     Select Case mvarGabimi
        Case 9
            MsgBox "Para se te hidhni notat ju duhet te zgjidhni nje nga opsionet e mesiperme.", vbInformation, "Hedhja e notave."
        Case 113
            MsgBox "Ju pasi zgjidhni semestrin duhet te zgjidhni dhe llojin e notes.", vbExclamation, "Hedhja e notave."
        Case 10
            MsgBox "Para se te hidhni notat per nje klase te caktuar ju duhet te jepni te dhenat identifikuese te kesaj klase " & Chr(10) & "qe jane numri i klases, indeksi si dhe viti shkollor", vbInformation, "Hedhja e notave."
        Case 111
            Dim vitFillimi As Integer
            vitFillimi = CInt(Mid(active_form.cboVitiShkollor.Text, 1, 4))
            If (vitFillimi <= 2006) Then
                MsgBox "Ju mund te hidhni notat e provimit te lirimit vetem nxenesve qe ndodhen ne klasen e 8-te.", vbInformation, "Hedhja e notave."
                active_form.cboKlasa.Text = "8"
            Else
                MsgBox "Ju mund te hidhni notat e provimit te lirimit vetem nxenesve qe ndodhen ne klasen e 9-te.", vbInformation, "Hedhja e notave."
                active_form.cboKlasa.Text = "9"
            End If
        Case 112
            MsgBox "Ju mund te hidhni notat e provimit te matures vete nxenesve qe ndodhen ne klasen e 12-te.", vbInformation, "Hedhja e notave."
            active_form.cboKlasa.Text = "12"
        
     End Select
End Sub
Private Sub menazhim_gabimi_frmModifikimeNota()
     Select Case mvarGabimi
      Case 11
         MsgBox " Ju duhet te zgjidhni nje nga opsionet e shkolles se mesme ose nëntëvjeçare ", vbInformation, "Modifikimi i notave."
      
      Case 1001
        MsgBox "Ju duhet te zgjidhni nje nga opsionet e llojit te notes qe doni te modifikoni.", vbExclamation, "Modifikimi i notave."

      Case 12
         MsgBox "Para se te konsultoni notat jepni vitin shkollor" + Chr(10) + "si edhe te paktën një nga të dhenat identifikuese per nxenesin " & Chr(10) & "qe mund te jene numri i amzes,emri, mbiemri ose atesia!", vbInformation, "Konsultimi i notave."
      Case Else

   End Select

End Sub
Private Sub menazhim_gabimi_frmStatistikaKlasa()
    
    Select Case mvarGabimi
       Case 13
          MsgBox "Ju duhet te jepni te dhenat e kerkuara para se te shfaqen klasat me statistikat perkatese ," & Chr(10) & "te  cilat jane semestri dhe viti shkollor.", vbInformation, "Statistika per klasat."
          
    End Select
End Sub
Private Sub menazhim_gabimi_frmStatistikaNxenes()
    Select Case mvarGabimi
       Case 14
          MsgBox "Ju duhet te jepni te dhenat e kerkuara para se te shfaqen klasat me statistikat perkatese , te  cilat jane semestri dhe viti shkollor.", vbInformation, "Statistika per nxenesit."
          
    End Select
End Sub
Private Sub menazhim_gabimi_frmInstrumenteKaloKlase()
    Select Case mvarGabimi
        Case 15
            MsgBox "Ju duhet te percaktoni cikli e shkolles para se te krueni ndonje veprim mbi klasat.", vbInformation, "Kalo nxenes."
        
    End Select
End Sub
Private Sub menazhim_gabimi_frmMesataretCikli()
    Select Case mvarGabimi
        Case 1
            MsgBox "Zgjidhni një prej semestreve për të shfaqur mesataret momentale sipas cikleve.", vbExclamation, "Mesataret momentale sipas cikleve"
    End Select
End Sub
Private Sub menazhim_gabimi_frmModifikimeGjeneralitete()
   Select Case mvarGabimi
      Case 17
         MsgBox " Ju duhet te zgjidhni nje nga opsionet e shkolles se mesme ose nëntëvjeçare ", vbInformation, "Modifikimi i gjeneraliteteve."

      Case 18
         MsgBox "Para se te konsultoni gjeneralitetet e nxënësit jepni te paktën" + Chr(10) + "një nga të dhënat identifikuese për nxënësin " & Chr(10) & "që mund të jenë numri i amzës,emri, mbiemri ose atësia!", vbInformation, "Konsultimi i notave."


      Case 19
         MsgBox " Ju duhet te jepni te gjitha te dhenat e kerkuara me lart para se te hidhni nje nxenes te ri ne amze.", vbInformation, "Modifikimi i gjeneraliteteve."

      Case 20
         MsgBox "Numri i amzes i dhene nga ju nuk eshte ne formatin e duhur.Ai duhet te permbaje vetem shifra.", vbInformation, "Modifikimi i gjeneraliteteve."
    Case 21
         MsgBox "Klasa e caktuar prej jush nuk i përket ciklit të nxënësit.", vbInformation, "Modifikimi i gjeneraliteteve."
   End Select
End Sub
Private Sub menazhim_gabimi_frmEleminoNxenes()
    Select Case mvarGabimi
        Case 21
            MsgBox " Ju duhet te zgjidhni nje nga opsionet e shkolles se mesme ose nëntëvjeçare ", vbInformation, "Eleminimi i nxenesit."
        Case 22
            MsgBox "Para se te largoni nxenesin jepni te dhenat identifikuese per nxenesin " & Chr(10) & "qe jane numri i amzes, emri,atesia ose mbiemri .", vbInformation, "Eleminimi i nxenesit."
    End Select
End Sub

Private Sub menazhim_gabimi_frmModifikoMungesa()
     Select Case mvarGabimi
      Case 1
         MsgBox "Duhet të zgjidhni më parë vitin shkollor, klasën dhe indeksin e klasës" + Chr(10) + "për të cilën do të modifikoni mungesat!", vbExclamation, "Modifikimi i mungesave."
      
      Case 2
        MsgBox "Data e fillimit duhet të jetë më e vogël ose e barabartë se data e mbarimit!", vbExclamation, "Modifikimi i mungesave."

      Case 3
         MsgBox "Intervali i dhënë i datave nuk përkon plotësisht me vitin shkollor të zgjedhur!", vbExclamation, "Modifikimi i mungesave"
      Case 4:
         MsgBox "Intervali i datave nuk duhet të jetë më i madh se 15 ditë!", vbExclamation, "Modifikimi i mungesave"
      Case 5:
         MsgBox "Ju nuk keni të drejta të modifikoni mungesat e klasës së zgjedhur!", vbExclamation, "Modifikimi i mungesave"
      Case Else

   End Select

End Sub

Private Sub menazhim_gabimi_frmNxenesitDalluar()
    Select Case mvarGabimi
      Case 1
         MsgBox "Duhet të zgjidhni më parë një prej cikleve e mesme ose nëntëvjeçare!", vbExclamation, "Nxënësit e dalluar."
      
      Case 2
        MsgBox "Duhet të zgjidhni më parë periudhën e vlerësimit për nxënësit" + Chr(10) + "Semestri i parë, Semestri i dytë ose Vjetore!", vbExclamation, "Nxënësit e dalluar."

      Case 3
         MsgBox "Duhet të zgjidhni më parë vitin shkollor!", vbExclamation, "Nxënësit e dalluar"
      Case 4:
         MsgBox "Duhet të zgjidhni të paktën një kriter për të përzgjedhur nxënësit e dalluar!", vbExclamation, "Nxënësit e dalluar"
      Case 5:
         MsgBox "Shkruani vlerën e mesatares kriter për nxënësit e dalluar!", vbExclamation, "Nxënësit e dalluar"
      Case 6:
         MsgBox "Mesatarja kriter për nxënësit e dalluar duhet të jetë numër me një shifër pas presjes dhjetore!", vbExclamation, "Nxënësit e dalluar"
      Case 7:
         MsgBox "Mesatarja kriter për nxënësit e dalluar nuk mund të jetë më e vogël se 5.0!", vbExclamation, "Nxënësit e dalluar"
      Case Else
   End Select
End Sub
Private Function kontrollo_numerAmze(nrAmza As String) As Boolean
  If nrAmza <> "" Then
    Dim p As String
    Dim ugjet As Boolean
    Dim k As Integer
    k = Len(nrAmza)
    Dim I As Integer
    I = 1
    Do While (I <= k) And Not ugjet
       p = Mid(nrAmza, I, 1)
       If Not (p = "1" Or p = "2" Or p = "3" Or p = "3" Or p = "4" Or p = "5" Or p = "6" Or p = "7" Or p = "8" Or p = "9" Or p = "0") Then
          ugjet = True
       End If
       I = I + 1
    Loop
    kontrollo_numerAmze = ugjet
       
  Else
    kontrollo_numerAmze = False
  End If
End Function

