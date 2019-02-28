VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusInfoProd"
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
' * Date             : 11/23/2004
' * Time             : 21:13
' * Module Name      : clsInfoProd
' * Module Filename  : clsInfoProd.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************


'local variable(s) to hold property value(s)
Private mvarErrorHandler As clsErrorHandler



Public Property Set ErrorHandler(ByVal vData As clsErrorHandler)
    Set mvarErrorHandler = vData
End Property


Public Property Get ErrorHandler() As clsErrorHandler
   Set ErrorHandler = mvarErrorHandler
End Property


Public Sub SetGUIColLenda_HedhjeGjeneraliteteVizualizoGride(objColLendet As colBusLenda)
   Dim objLenda         As New clsBusLenda
   Dim objGridHandler   As New clsGridHandler
   Dim p                As Double

   With active_form

      p = (1 / (3 + objColLendet.Count)) * 100

      .SGGrid1.Columns.RemoveAll False
      objGridHandler.AddCol .SGGrid1, "Klasa", "Klasa", p
      objGridHandler.AddCol .SGGrid1, "Viti Shkollor", "Viti Shkollor", p

      For Each objLenda In objColLendet

         objGridHandler.AddCol .SGGrid1, objLenda.Key, objLenda.Emri, p, , sgtDouble

      Next objLenda

      objGridHandler.AddCol .SGGrid1, "Vrejtje", "Vrejtje", p

      objGridHandler.AddCol .SGGrid2, "Matematike", "Matematike", 25, , sgtDouble
      objGridHandler.AddCol .SGGrid2, "Letersi", "Letersi", 25, , sgtDouble
      objGridHandler.AddCol .SGGrid2, "Histori", "Histori", 25, , sgtDouble

   End With

   Set objGridHandler = Nothing
   Set objLenda = Nothing
End Sub



Public Sub SetGUIColLenda_(objColLenda As colBusLenda)

  
End Sub


