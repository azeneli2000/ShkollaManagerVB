VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colBusLenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsLenda"
Attribute VB_Ext_KEY = "Member0" ,"clsLenda"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'local variable to hold collection
Private mCol As Collection


Public Property Get Item(vntIndexKey As Variant) As clsBusLenda
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


Public Function Add(objNewMember As clsBusLenda, Optional sKey As String) As clsBusLenda

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
   Dim oItem            As clsBusLenda
   Set oItem = mCol(CStr(vIndexKey))

   'If referenced ok - no error
   Exists = True
   Exit Function

KeyDoesNotExist:
   Err.Clear
   Exists = False

End Function


Public Function Clone() As colBusLenda
   Dim objLenda         As clsBusLenda
   Dim newCol           As New colBusLenda

   For Each objLenda In mCol
      newCol.Add objLenda.Clone
   Next objLenda

   Set Clone = newCol
   Set newCol = Nothing

End Function


Public Function Load(rs As ADODB.Recordset) As colBusLenda
   Dim oLenda           As clsBusLenda

   If rs Is Nothing Then
      Exit Function
   End If

   While Not rs.EOF
      Set oLenda = New clsBusLenda

      oLenda.Emri = rs("LN_EMRI")
      oLenda.VitiShkollor = rs("LN_VITISHKOLLOR")
      oLenda.Key = Str(rs("LN_ID"))

      mCol.Add oLenda

      rs.MoveNext
   Wend

   Set Load = Me

End Function


Public Function LoadV(vParams As Variant) As colBusLenda
   Dim oLenda           As clsBusLenda

   If Not IsArray(vParams) Then
      Exit Function
   End If

   For i = 0 To UBound(vParams, 2)
      Set oLenda = New clsBusLenda

      oLenda.Emri = vParams(1, i)
      oLenda.VitiShkollor = vParams(2, i)
      oLenda.Key = vParams(0, i)

      mCol.Add oLenda

   Next i

   Set LoadV = Me

End Function


Public Function GetGUIColLenda_() As colBusLenda
   Dim objLenda         As clsBusLenda
   '' to do
   With active_form

      mCol.Add objLenda
   End With

   Set GetGUIColLenda = Me

   Set objLenda = Nothing

End Function
