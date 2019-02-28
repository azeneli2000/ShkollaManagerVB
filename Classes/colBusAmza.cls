VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colBusAmza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsAmza"
Attribute VB_Ext_KEY = "Member0" ,"clsAmza"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
' * Author           :
' * Web Site         :
' * E-Mail           :
' * Date             : 11/23/2004
' * Time             : 21:23
' * Module Name      : colBusAmza
' * Module Filename  : colBusAmza.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************



'local variable to hold collection
Private mCol As Collection


Public Property Get Item(vntIndexKey As Variant) As clsBusAmza
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


Public Function Add(objNewMember As clsBusAmza, Optional sKey As String) As clsBusAmza

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
   Dim oItem            As clsBusAmza
   Set oItem = mCol(CStr(vIndexKey))

   'If referenced ok - no error
   Exists = True
   Exit Function

KeyDoesNotExist:
   Err.Clear
   Exists = False

End Function


Public Function Clone() As colBusAmza
   Dim objAmza          As clsBusAmza
   Dim newCol           As New colBusAmza

   For Each objAmza In mCol
      newCol.Add objAmza.Clone
   Next objAmza

   Set Clone = newCol
   Set newCol = Nothing

End Function


Public Function Load(rs As ADODB.Recordset) As colBusAmza
   Dim objAmza          As clsBusAmza

   While Not rs.EOF
      Set objAmza = New clsBusAmza

      With objAmza
         .Key = rs("AMZ_ID")
         .No = rs("AMZ_NO")
         .Klasa = rs("AMZ_KLASA")
         .KlaseIndex = rs("AMZ_KLASEINDEX")
         .Emri = rs("AMZ_EMRI")
         .Mbiemri = rs("AMZ_MBIEMRI")
         .Atesia = rs("AMZ_ATESIA")
         .Seksi = rs("AMZ_SEKSI")
         .Memesia = rs("AMZ_MEMESIA")
         .Datelindja = rs("AMZ_DATELINDJA")
         .Vendlindja = rs("AMZ_VENDLINDJA")
      End With
      mCol.Add objAmza

      rs.MoveNext
   Wend
   
   Set Load = Me

End Function


Public Function LoadV(vParams() As Variant) As colBusAmza
   Dim objAmza          As clsBusAmza

   If Not IsArray(vParams) Then
      Exit Function
   End If

   For i = 0 To UBound(vParams, 2)
      Set objAmza = New clsBusAmza

      With objAmza
         .Key = vParams(0, i)
         .No = vParams(1, i)
         .Klasa = vParams(2, i)
         .Emri = vParams(3, i)
         .Mbiemri = vParams(4, i)
         .Atesia = vParams(5, i)
         .Memesia = vParams(6, i)
         .Seksi = vParams(7, i)
         .Vendlindja = vParams(8, i)
         .Datelindja = vParams(9, i)
         .Vrejtje = vParams(10, i)
         .KlaseIndex = vParams(11, i)
      End With

      mCol.Add objAmza
   Next i

   Set LoadV = Me

End Function


Public Function GetGUIColAmza_() As colBusAmza
   Dim objAmza          As clsBusAmza
   '' to do
   With active_form

      Me.Add objAmza
   End With

   Set GetGUIColAmza_ = Me
   Set objAmza = Nothing
End Function
