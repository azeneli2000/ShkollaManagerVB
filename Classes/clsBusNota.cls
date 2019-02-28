VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusNota"
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
' * Time             : 20:54
' * Module Name      : clsNota
' * Module Filename  : clsNota.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************



'local variable(s) to hold property value(s)
Private mvarNotat As String
Private mvarData As Date

Public Enum NOTE_TIP_ENUM
    ntMomentaleI = 1
    ntMomentaleII = 2
    ntSemestraleI = 3
    ntSemestraleII = 4
    ntVjetore = 5
    ntMungese = 6
    ntDetyreKontrolli = 7
End Enum

Private mvarKlasa As String
Private mvarKlaseIndex As String
Private mvarErrorHandler As New clsErrorHandler
Private mvarLendaKey As String
Private mvarAmzKey As String
Private mvarKey As String
Private mvarNoteType As NOTE_TIP_ENUM



Public Property Let NoteType(ByVal vData As NOTE_TIP_ENUM)
    mvarNoteType = vData
End Property


Public Property Get NoteType() As NOTE_TIP_ENUM
   NoteType = mvarNoteType
End Property


Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property


Public Property Get Key() As String
   Key = mvarKey
End Property


Public Property Let AmzKey(ByVal vData As String)
    mvarAmzKey = vData
End Property


Public Property Get AmzKey() As String
   AmzKey = mvarAmzKey
End Property


Public Property Let LendaKey(ByVal vData As String)
    mvarLendaKey = vData
End Property


Public Property Get LendaKey() As String
   LendaKey = mvarLendaKey
End Property


Public Property Set ErrorHandler(ByVal vData As clsErrorHandler)
    Set mvarErrorHandler = vData
End Property


Public Property Get ErrorHandler() As clsErrorHandler
   Set ErrorHandler = mvarErrorHandler
End Property


Public Property Let KlaseIndex(ByVal vData As String)
    mvarKlaseIndex = vData
End Property


Public Property Get KlaseIndex() As String
   KlaseIndex = mvarKlaseIndex
End Property


Public Property Let Klasa(ByVal vData As String)
    mvarKlasa = vData
End Property


Public Property Get Klasa() As String
   Klasa = mvarKlasa
End Property


Public Property Let Data(ByVal vData As Date)
    mvarData = vData
End Property


Public Property Get Data() As Date
   Data = mvarData
End Property


Public Property Let Notat(ByRef vData As String)
    mvarNotat = vData
End Property


Public Property Get Notat() As String
   Notat = mvarNotat
End Property


Private Sub Class_Terminate()
   Set ErrorHandler = Nothing
End Sub


Public Function Clone() As clsBusNota
   Dim ClonedObject     As New clsBusNota

   With ClonedObject
      .Notat = Me.Notat
      .Data = Me.Data
      .Klasa = Me.Klasa
      .KlaseIndex = Me.KlaseIndex
      .LendaKey = Me.LendaKey
      .NoteType = Me.NoteType
      Set .ErrorHandler = Me.ErrorHandler.Clone
   End With

   Set Clone = ClonedObject

   Set ClonedObject = Nothing
End Function


Public Function Load(rs As ADODB.Recordset) As clsBusNota

   If Not rs.EOF Then
      With Me
         .Key = rs("NT_ID")
         .AmzKey = rs("NT_NTAID")
         .LendaKey = rs("LN_ID")
         .Notat = rs("NT_VLERESIMI")
         .Data = rs("NT_DATA")
         .NoteType = Switch(rs("NT_MOMENTALES1") = YES, ntMomentaleI, rs("NT_MOMENTALES2"), ntMomentaleII, rs("NT_SEMESTRI1"), ntSemestraleI, rs("NT_SEMESTRI2") = YES, ntSemestraleII, rs("NT_VJETORE") = YES, ntVjetore, rs("NT_MUNGESE") = YES, ntMungese, rs("DETYREKONTROLLI") = YES, ntDetyreKontrolli)

      End With
   End If

   Set Load = Me

End Function


Public Function LoadV(vParams() As Variant) As clsBusNota

   If Not IsArray(vParams) Then
      Exit Function
   End If

   With Me
      .Key = vParams(0, 0)
      .AmzKey = vParams(1, 0)
      .LendaKey = vParams(8, 0)
      .Notat = vParams(2, 0)
      .Data = vParams(3, 0)
      .NoteType = Switch(vParams(4, 0), ntMomentaleI, vParams(9, 0), ntMomentaleII, vParams(5, 0), ntSemestraleI, vParams(6, 0), ntSemestraleII, vParams(7, 0), ntVjetore, vParams(10, 0), ntMungese, vParams(11, 0), ntDetyreKontrolli)
   End With

   Set LoadV = Me

End Function


Public Function GetArray() As Variant

   With Me
      GetArray = Array(.Key, .AmzKey, .Notat, .Data, IIf(.NoteType = ntMomentaleI, "Yes", "No"), IIf(.NoteType = ntSemestraleI, "Yes", "No"), IIf(.NoteType = ntSemestraleII, "Yes", "No"), IIf(.NoteType = ntVjetore, "Yes", "No"), .LendaKey, IIf(.NoteType = ntMomentaleII, "Yes", "No"), IIf(.NoteType = ntMungese, "Yes", "No"), IIf(.NoteType = ntDetyreKontrolli, "Yes", "No"))
   End With

End Function


Public Function GetGUINotat_() As clsBusNota

   ' to do
   With active_form

   End With

   Set GetGUINotat_HedhjeGjeneralitete = Me

End Function
