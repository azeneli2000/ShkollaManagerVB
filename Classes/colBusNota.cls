VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colBusNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsNota"
Attribute VB_Ext_KEY = "Member0" ,"clsNota"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
' * Author           :
' * Web Site         :
' * E-Mail           :
' * Date             : 11/23/2004
' * Time             : 21:24
' * Module Name      : colBusNota
' * Module Filename  : colBusNota.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************




'local variable to hold collection
Private mCol As Collection


Public Property Get Item(vntIndexKey As Variant) As clsBusNota
Attribute Item.VB_UserMemId = 0
   Set Item = mCol(vntIndexKey)
End Property


Public Property Get Count() As Long
   Count = mCol.Count
End Property


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
   Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
   Set mCol = Nothing
End Sub


Public Sub Clear()
   Set mCol = New Collection
End Sub


Public Sub Remove(vntIndexKey As Variant)
   mCol.Remove vntIndexKey
End Sub


Public Function Add(objNewMember As clsBusNota, Optional sKey As String) As clsBusNota

   If Len(sKey) = 0 Then
      mCol.Add objNewMember
   Else
      mCol.Add objNewMember, sKey
   End If

   'return the object created
   Set Add = objNewMember
   Set objNewMember = Nothing

End Function


Public Function Exists(vIndexKey As Variant) As Boolean

   'Error handling
   On Error GoTo KeyDoesNotExist

   'Attempt to reference item
   Dim oItem            As clsBusNota
   Set oItem = mCol(CStr(vIndexKey))

   'If referenced ok - no error
   Exists = True
   Exit Function

KeyDoesNotExist:
   Err.Clear
   Exists = False

End Function


Public Function Clone() As colBusNota
   Dim objNota          As clsBusNota
   Dim newCol           As New colBusNota

   For Each objNota In mCol
      newCol.Add objNota.Clone
   Next objNota

   Set Clone = newCol
   Set newCol = Nothing

End Function


Public Function Load(rs As ADODB.Recordset) As colBusNota
   Dim objNota          As clsBusNota

   While Not rs.EOF
      Set objNota = New clsBusNota

      With objNota
         .Key = rs("NT_ID")
         .AmzKey = rs("NT_NTAID")
         .LendaKey = rs("LN_ID")
         .Notat = rs("NT_VLERESIMI")
         .Data = rs("NT_DATA")
         .NoteType = Switch(rs("NT_MOMENTALES1") = YES, ntMomentaleI, rs("NT_MOMENTALES2"), ntMomentaleII, rs("NT_SEMESTRI1"), ntSemestraleI, rs("NT_SEMESTRI2") = YES, ntSemestraleII, rs("NT_VJETORE") = YES, ntVjetore, rs("NT_MUNGESE") = YES, ntMungese, rs("DETYREKONTROLLI") = YES, ntDetyreKontrolli)

      End With

      mCol.Add objNota
      rs.MoveNext

   Wend

   Set Load = Me
End Function


Public Function LoadV(vParams() As Variant) As colBusNota
   Dim objNota          As clsBusNota

   If Not IsArray(vParams) Then
      Exit Function
   End If

   For i = 0 To UBound(vParams, 2)
      Set objNota = New clsBusNota

      With objNota
         .Key = vParams(0, i)
         .AmzKey = vParams(1, i)
         .LendaKey = vParams(8, i)
         .Notat = vParams(2, i)
         .Data = vParams(3, i)
         .NoteType = Switch(vParams(4, i), ntMomentaleI, vParams(9, i), ntMomentaleII, vParams(5, i), ntSemestraleI, vParams(6, i), ntSemestraleII, vParams(7, i), ntVjetore, vParams(10, i), ntMungese, vParams(11, i), ntDetyreKontrolli)
      End With

      mCol.Add objNota

   Next i

   Set LoadV = Me

End Function


Public Function GetGUINotat_HedhjeGjeneraliteteRuaj() As colBusNota
   Dim objNota          As clsBusNota

   With active_form
      Dim arr              As SGArray

      Set arr = .SGGrid1.Array

      For i = 3 To .SGGrid1.DataColCount - 2
         Set objNota = New clsBusNota
         objNota.Klasa = arr(0, 0)
         objNota.KlaseIndex = arr(0, 0)
         objNota.LendaKey = .SGGrid1.Columns(i).Key
         objNota.Notat = arr(0, i)
         objNota.Data = Now
         objNota.NoteType = ntVjetore

         Me.Add objNota
      Next i
   End With
   
   Set GetGUINotat_HedhjeGjeneraliteteRuaj = Me
   Set objNota = Nothing

End Function


Public Function GetGUINotat_() As colBusNota
   Dim objNota          As clsBusNota

   '' to do
   With active_form

      Me.Add objNota

   End With
   Set GetGUINotat_ = Me
   Set objNota = Nothing
End Function
