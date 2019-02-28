Attribute VB_Name = "modAPILibrary"

' * Author           :
' * Web Site         :
' * E-Mail           :
' * Date             : 11/18/2004
' * Time             : 15:35
' * Module Name      : modAPILibrary
' * Module Filename  : modAPILibrary.bas
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************


Option Explicit



Public hWnd As Long

Global Const HH_DISPLAY_TOPIC = &H0
Global Const HH_SET_WIN_TYPE = &H4
Global Const HH_GET_WIN_TYPE = &H5
Global Const HH_GET_WIN_HANDLE = &H6
Global Const HH_DISPLAY_TEXT_POPUP = &HE      ' Display string resource ID or
                                        ' text in a pop-up window.
Global Const HH_HELP_CONTEXT = &HF            ' Display mapped numeric value in
                                        ' dwData.
Global Const HH_TP_HELP_CONTEXTMENU = &H10    ' Text pop-up help, similar to
                                        ' WinHelp's HELP_CONTEXTMENU.
Global Const HH_TP_HELP_WM_HELP = &H11        ' text pop-up help, similar to
                                        ' WinHelp's HELP_WM_HELP.


Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long


Public Sub CallHelp(contextId As Long)
   Dim hwndHelp         As Long
   'The return value is the window handle of the created help window.
   hwndHelp = HtmlHelp(hWnd, App.Path & "\Help\Help.chm", HH_HELP_CONTEXT, contextId)
End Sub
