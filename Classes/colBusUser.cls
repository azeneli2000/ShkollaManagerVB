VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colBusUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsUser"
Attribute VB_Ext_KEY = "Member0" ,"clsUser"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
' * Author           :
' * Web Site         :
' * E-Mail           :
' * Date             : 11/23/2004
' * Time             : 21:21
' * Module Name      : colBusUser
' * Module Filename  : colBusUser.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************





'local variable to hold collection
Private mCol As Collection


Public Property Get Item(vntIndexKey As Variant) As clsBusUser
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

Public Function Add(objNewMember As clsBusUser, Optional sKey As String) As clsBusUser

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
   Dim oItem            As clsBusUser
   Set oItem = mCol(CStr(vIndexKey))

   'If referenced ok - no error
   Exists = True
   Exit Function

KeyDoesNotExist:
   Err.Clear
   Exists = False

End Function


Public Function Clone() As colBusUser
   Dim objUser          As clsBusUser
   Dim ClonedObject     As New colBusUser

   For Each objUser In mCol
      ClonedObject.Add objUser.Clone
   Next objUser

   Set Clone = ClonedObject

   Set objUser = Nothing
   Set ClonedObject = Nothing

End Function


Public Function Load(rs As ADODB.Recordset) As colBusUser
   Dim objUser As clsBusUser

   While Not rs.EOF
      Set objUser = New clsBusUser
      With objUser
         .Key = rs("USER_ID")
         .UserName = rs("USER_NAME")
         .UserPassword = rs("USER_PASSWORD")
      End With

      mCol.Add objUser

      rs.MoveNext
   Wend

   Set Load = Me

End Function


Public Function LoadV(vParams As Variant) As colBusUser
   Dim objUser          As clsBusUser

   For i = LBound(vParams) To UBound(vParams)
      Set objUser = New clsBusUser
      With objUser
         .Key = vParams(0, 0)
         .UserName = vParams(1, 0)
         .UserPassword = vParams(2, 0)
      End With

      mCol.Add objUser

   Next i

   Set LoadV = Me

End Function


