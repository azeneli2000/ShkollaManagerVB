VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colInfoProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsInfoProd"
Attribute VB_Ext_KEY = "Member0" ,"clsInfoProd"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
' * Author           :
' * Web Site         :
' * E-Mail           :
' * Date             : 11/22/2004
' * Time             : 22:58
' * Module Name      : colInfoProd
' * Module Filename  :
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************


'local variable to hold collection
Private mCol As Collection

Public Function Add(objNewMember As clsInfoProd, Optional sKey As String) As clsInfoProd
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsInfoProd
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

' Kthen true nqs elementi me index vIndexKey ekzizton
' kujdes indeksi fillon nga numri 1
Public Function Exists(vIndexKey As Variant) As Boolean

   'Error handling
   On Error GoTo KeyDoesNotExist

   'Attempt to reference item
   Dim oItem As clsInfoProd
   Set oItem = mCol(CStr(vIndexKey))

   'If referenced ok - no error
   Exists = True
   Exit Function

KeyDoesNotExist:
   Err.Clear
   Exists = False

End Function

Public Sub Clear()
  Set mCol = New Collection
End Sub

Public Function Clone() As colInfoProd
  Dim objInfoProd As clsInfoProd
  Dim newCol As New Collection
  
  For Each objAmza In mCol
      newCol.Add objInfoProd.Clone
  Next objAmza
  
  Set Clone = newCol
  Set newCol = Nothing
  
End Function



