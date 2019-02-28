VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusNotaAmza"
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
' * Date             : 11/24/2004
' * Time             : 16:17
' * Module Name      : clsBusNotaAmza
' * Module Filename  :
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************





'local variable(s) to hold property value(s)
Private mvarAmzaKey As String
Private mvarLendaKey As String
Private mvarNotaP1 As Double
Private mvarNotaP2 As Double
Private mvarNotaP3 As Double
Private mvarNotaP4 As Double
Private mvarKey As String



Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property


Public Property Get Key() As String
   Key = mvarKey
End Property


Public Property Let NotaP4(ByVal vData As Double)
    mvarNotaP4 = vData
End Property


Public Property Get NotaP4() As Double
   NotaP4 = mvarNotaP4
End Property


Public Property Let NotaP3(ByVal vData As Double)
    mvarNotaP3 = vData
End Property


Public Property Get NotaP3() As Double
   NotaP3 = mvarNotaP3
End Property


Public Property Let NotaP2(ByVal vData As Double)
    mvarNotaP2 = vData
End Property


Public Property Get NotaP2() As Double
   NotaP2 = mvarNotaP2
End Property


Public Property Let NotaP1(ByVal vData As Double)
    mvarNotaP1 = vData
End Property


Public Property Get NotaP1() As Double
   NotaP1 = mvarNotaP1
End Property


Public Property Let LendaKey(ByVal vData As String)
    mvarLendaKey = vData
End Property


Public Property Get LendaKey() As String
   LendaKey = mvarLendaKey
End Property


Public Property Let AmzaKey(ByVal vData As String)
    mvarAmzaKey = vData
End Property


Public Property Get AmzaKey() As String
   AmzaKey = mvarAmzaKey
End Property


Public Function Clone() As clsBusNotaAmza
   Dim objClone         As New clsBusNotaAmza

   With objClone
      .AmzaKey = Me.AmzaKey
      .Key = Me.Key
      .LendaKey = Me.LendaKey
      .NotaP1 = Me.NotaP1
      .NotaP2 = Me.NotaP2
      .NotaP3 = Me.NotaP3
      .NotaP4 = Me.NotaP4
   End With

   Set Clone = objClone

   Set objClone = Nothing

End Function


Public Function Load(rs As ADODB.Recordset) As clsBusNotaAmza

   If Not rs.EOF Then
      With Me
         .AmzaKey = rs("NTA_AMZID")
         .Key = rs("NTA_ID")
         .LendaKey = rs("NTA_LNID")
         .NotaP1 = rs("NTA_NOTAP1")
      End With

   End If

   Set Load = Me

End Function


Public Function LoadV(vParams As Variant) As clsBusNotaAmza

   With Me
      .AmzaKey = vParams(0, 0)
      .Key = vParams(5, 0)
      .LendaKey = vParams(6, 0)
      .NotaP1 = vParams(1, 0)
   End With

   Set LoadV = Me
End Function


Public Function GetArray() As Variant

   With Me
      GetArray = Array(.AmzaKey, .NotaP1, .NotaP2, .NotaP3, .NotaP4, .Key, .LendaKey)
   End With

End Function


Public Function GetGUINotaAmza_() As clsBusNotaAmza

   ' to do
   With active_form

   End With

   Set GetGUINotaAmza_ = Me
End Function
