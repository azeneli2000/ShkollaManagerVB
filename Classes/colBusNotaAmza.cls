VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colBusNotaAmza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsBusNotaAmza"
Attribute VB_Ext_KEY = "Member0" ,"clsBusNotaAmza"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
' * Author           :
' * Web Site         :
' * E-Mail           :
' * Date             : 11/24/2004
' * Time             : 16:37
' * Module Name      : colBusNotaAmz
' * Module Filename  :
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************


'local variable to hold collection
Private mCol As Collection


Public Property Get Item(vntIndexKey As Variant) As clsBusNotaAmza
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


Public Function Add(objNewMember As clsBusNotaAmza, Optional sKey As String) As clsBusNotaAmza

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
   Dim oItem            As clsBusNotaAmza
   Set oItem = mCol(CStr(vIndexKey))

   'If referenced ok - no error
   Exists = True
   Exit Function

KeyDoesNotExist:
   Err.Clear
   Exists = False

End Function


Public Function Clone() As colBusNotaAmza
   Dim objNotaAmza      As clsBusNotaAmza
   Dim ClonedObject     As New colBusNotaAmza

   For Each objNotaAmza In mCol
      ClonedObject.Add objNotaAmza.Clone
   Next objNotaAmza

   Set Clone = ClonedObject

   Set ClonedObject = Nothing
   Set objNotaAmza = Nothing
End Function


Public Function Load(rs As ADODB.Recordset) As colBusNotaAmza
   Dim objNotaAmza      As clsBusNotaAmza

   While Not rs.EOF
      Set objNotaAmza = New clsBusNotaAmza

      With objNotaAmza
         .AmzaKey = rs("NTA_AMZID")
         .Key = rs("NTA_ID")
         .LendaKey = rs("NTA_LNID")
         .NotaP1 = rs("NTA_NOTAP1")
      End With

      mCol.Add objNotaAmza

      rs.MoveNext
   Wend

   Set Load = Me

End Function


Public Function LoadV(vParams As Variant) As colBusNotaAmza
   Dim objNotaAmza      As clsBusNotaAmza

   For i = LBound(vParams) To UBound(vParams)
      Set objNotaAmza = New clsBusNotaAmza

      With objNotaAmza
         .AmzaKey = vParams(0, i)
         .Key = vParams(5, i)
         .LendaKey = vParams(6, i)
         .NotaP1 = vParams(1, i)
      End With

      mCol.Add objNotaAmza

   Next i

   Set LoadV = Me
End Function


Public Function GetGUINotatAmza_HedhjeGjeneraliteteRuaj() As colBusNotaAmza
   Dim objNotaAmza      As clsBusNotaAmza

   With active_form
      Set arr = .SGGrid2.Array

      For i = 0 To .SGGrid2.DataColCount - 1
         Set objNotaAmza = New clsBusNotaAmza
         objNotaAmza.LendaKey = .SGGrid1.Columns(i).Key
         objNotaAmza.NotaP1 = arr(0, 0)
         objNotaAmza.NotaP2 = arr(0, 1)

         Me.Add objNotaAmza
      Next i
   End With

   Set GetGUINotatAmza_HedhjeGjeneraliteteRuaj = Me
   Set objNotaAmza = Nothing
End Function


Public Function GetGUINotatAmza_() As colBusNotaAmza
   Dim objNotaAmza      As clsBusNotaAmza

   '' to do
   With active_form

      Me.Add objNotaAmza

   End With

   Set GetGUINotatAmza_ = Me
   Set objNotaAmza = Nothing
End Function
