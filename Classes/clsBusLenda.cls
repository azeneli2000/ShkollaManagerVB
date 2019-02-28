VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusLenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

' * Author           :
' * Web Site         :
' * E-Mail           :
' * Date             : 11/22/2004
' * Time             : 15:44
' * Module Name      : clsLenda
' * Module Filename  : clsLenda.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************

'local variable(s) to hold property value(s)
Private mvarEmri As String
Private mvarVitiShkollor As String
Private mvarKlasa As Variant
Private mvarKey As String




Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property


Public Property Get Key() As String
   Key = mvarKey
End Property


Public Property Let Klasa(ByVal vData As Variant)
    mvarKlasa = vData
End Property


Public Property Set Klasa(ByVal vData As Variant)
    Set mvarKlasa = vData
End Property


Public Property Get Klasa() As Variant
   If IsObject(mvarKlasa) Then
      Set Klasa = mvarKlasa
   Else
      Klasa = mvarKlasa
   End If
End Property


Public Property Let VitiShkollor(ByVal vData As String)
    mvarVitiShkollor = vData
End Property


Public Property Get VitiShkollor() As String
   VitiShkollor = mvarVitiShkollor
End Property


Public Property Let Emri(ByVal vData As String)
    mvarEmri = vData
End Property


Public Property Get Emri() As String
   Emri = mvarEmri
End Property


Public Function Clone() As clsBusLenda
   Dim ClonedObject     As New clsBusLenda

   With ClonedObject
      .Emri = Me.Emri
      .Klasa = Me.Klasa
      .VitiShkollor = Me.VitiShkollor
      .Key = Me.Key
   End With

   Set Clone = ClonedObject

   Set ClonedObject = Nothing
End Function


Public Function Load(rs As ADODB.Recordset) As clsBusLenda
   
   If rs Is Nothing Then
      Exit Function
   End If

   If Not rs.EOF Then
      With Me
         .Emri = rs("LN_EMRI")
         .VitiShkollor = rs("LN_VITISHKOLLOR")
         .Key = Str(rs("LN_ID"))
      End With
   End If

   Set Load = Me

End Function


Public Function LoadV(vParams As Variant) As clsBusLenda

   If Not IsArray(vParams) Then
      Exit Function
   End If

   With Me
      .Emri = vParams(1, 0)
      .VitiShkollor = vParams(2, 0)
      .Key = vParams(0, 0)
   End With

   Set LoadV = Me

End Function

Public Function GetArray() As Variant

   With Me
      GetArray = Array(.Key, .Emri, .VitiShkollor)
   End With

End Function


Public Function GetGUILende_() As clsBusLenda

   ' to do
   With active_form

   End With

   Set GetGUILende_ = Me

End Function
