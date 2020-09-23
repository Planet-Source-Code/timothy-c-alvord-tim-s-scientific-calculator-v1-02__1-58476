Attribute VB_Name = "ColorButtons"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2004 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'***********************************************************************
'*                                                                     *
'*  Copyright 1999 by Steve Derderian - The National Software Company  *
'*                                                                     *
'*  Overview:                                                          *
'*                                                                     *
'*  This module allows you to colour the button text in a Visual Basic *
'*  application.  The module can only be used from VB program code     *
'*  (probably in the Form_Load event).                                 *
'*                                                                     *
'*  Button text will appear black in the development environment.  It  *
'*  will only be coloured while the program is running. All other      *
'*  button properties, methods and events will work normally.          *
'*                                                                     *
'*                                                                     *
'*  Only three steps are required to use this module.                  *
'*      1.  Include this module in your VB project.                    *
'*      2.  When you add a button to the form set the style property   *
'*          to Graphical.                                              *
'*      3.  Call "RegisterButton" for each button you want to colour.  *
'*                                                                     *
'*                                                                     *
'*                                                                     *
'*                                                                     *
'*  RegisterButton:                                                    *
'*                                                                     *
'*  Used to start colouring a button's text.                           *
'*                                                                     *
'*  Syntax -- RegisterButton(<Button>, <Forecolor>)                    *
'*                                                                     *
'*  Part               Description                                     *
'*  ------------------------------------------------------------       *
'*  Button             The command button to register                  *
'*  Forecolor          The colour for the button text                  *
'*                                                                     *
'*  Returned Value -- Returns a Boolean value.  True if the            *
'*  registration succeeded and false if it failed.                     *
'*                                                                     *
'*  Remarks -- To change the button text colour, call RegisterButton   *
'*  again with the new colour.  This will not register the button      *
'*  twice.  It will only change the colour of an already registered    *
'*  button.                                                            *
'*                                                                     *
'*                                                                     *
'*                                                                     *
'*                                                                     *
'*  UnregisterButton:                                                  *
'*                                                                     *
'*  Used to stop colouring a button's text.                            *
'*                                                                     *
'*  Syntax -- UnregisterButton(<Button>)                               *
'*                                                                     *
'*  Part               Description                                     *
'*  ------------------------------------------------------------       *
'*  Button             The command button to unregister                *
'*                                                                     *
'*  Return Value -- Returns a Boolean value.  True if the              *
'*  unregistration succeeded and false if it failed.                   *
'*                                                                     *
'*  Remarks -- You don't need to unregister all button that were       *
'*  registered.  This will automatically be done when a form is        *
'*  closed.  This function is only provided so that a VB program may   *
'*  stop colouring a button before the form is closed.                 *
'*                                                                     *
'***********************************************************************

Private colButtons  As New Collection
Private Const KeyConst = "K"
Private Const PROP_COLOR = "SMDColor"
Private Const PROP_HWNDPARENT = "SMDhWndParent"
Private Const PROP_LPWNDPROC = "SMDlpWndProc"
Private Const GWL_WNDPROC As Long = (-4)
Private Const ODA_SELECT As Long = &H2
Private Const ODS_SELECTED As Long = &H1
Private Const ODS_FOCUS As Long = &H10
Private Const ODS_BUTTONDOWN As Long = ODS_FOCUS Or ODS_SELECTED
Private Const WM_DESTROY As Long = &H2
Private Const WM_DRAWITEM As Long = &H2B
Private Const VER_PLATFORM_WIN32_NT As Long = 2

Private Type RECT
   Left        As Long
   Top         As Long
   Right       As Long
   Bottom      As Long
End Type

Private Type SIZE
   cx          As Long
   cy          As Long
End Type

Private Type DRAWITEMSTRUCT
   CtlType     As Long
   CtlID       As Long
   itemID      As Long
   itemAction  As Long
   itemState   As Long
   hWndItem    As Long
   hDC         As Long
   rcItem      As RECT
   itemData    As Long
End Type

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Declare Function CallWindowProc Lib "user32" _
    Alias "CallWindowProcA" _
   (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal msg As Long, _
    ByVal wParam As Long, _
    lParam As DRAWITEMSTRUCT) As Long

Private Declare Function GetParent Lib "user32" _
    (ByVal hWnd As Long) As Long

Private Declare Function GetProp Lib "user32" _
    Alias "GetPropA" _
   (ByVal hWnd As Long, _
    ByVal lpString As String) As Long

Private Declare Function GetTextExtentPoint32 Lib "gdi32" _
    Alias "GetTextExtentPoint32A" _
   (ByVal hDC As Long, _
    ByVal lpSz As String, _
    ByVal cbString As Long, _
    lpSize As SIZE) As Long

Private Declare Function RemoveProp Lib "user32" _
    Alias "RemovePropA" _
   (ByVal hWnd As Long, _
    ByVal lpString As String) As Long

Private Declare Function SetProp Lib "user32" _
    Alias "SetPropA" _
   (ByVal hWnd As Long, _
    ByVal lpString As String, _
    ByVal hData As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" _
    (ByVal hDC As Long, _
    ByVal crColor As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" _
   (ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare Function TextOut Lib "gdi32" _
    Alias "TextOutA" _
   (ByVal hDC As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal lpString As String, _
    ByVal nCount As Long) As Long
    
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
  (lpVersionInformation As Any) As Long
    


Private Function FindButton(sKey As String) As Boolean

   Dim cmdButton As CommandButton
   
   On Error Resume Next
   Set cmdButton = colButtons.Item(sKey)
   FindButton = (Err.Number = 0)

End Function


Private Function GetKey(hWnd As Long) As String

   GetKey = KeyConst & hWnd

End Function


Private Function ProcessButton(ByVal hWnd As Long, _
                               ByVal uMsg As Long, _
                               ByVal wParam As Long, _
                               lParam As DRAWITEMSTRUCT, _
                               sKey As String) As Long

   Dim cmdButton       As CommandButton
   Dim bRC             As Boolean
   Dim lRC             As Long
   Dim x               As Long
   Dim y               As Long
   Dim lpWndProC       As Long
   Dim lButtonWidth    As Long
   Dim lButtonHeight   As Long
   Dim lPrevColor      As Long
   Dim lColor          As Long
   Dim TextSize        As SIZE
   Dim sCaption        As String
   
   Const PushOffset = 2
   
   Set cmdButton = colButtons.Item(sKey)
   sCaption = cmdButton.Caption
   
   lColor = GetProp(cmdButton.hWnd, PROP_COLOR)
   lPrevColor = SetTextColor(lParam.hDC, lColor)
   
  'in Pixels/Logical Units
   lRC = GetTextExtentPoint32(lParam.hDC, sCaption, Len(sCaption), TextSize)
   
  'in Pixels/Logical Units
   lButtonHeight = lParam.rcItem.Bottom - lParam.rcItem.Top
   lButtonWidth = lParam.rcItem.Right - lParam.rcItem.Left
   
  'the button is pressed! Offset the text
  'so it looks like the button is pushed
    If ((lParam.itemState And ODS_BUTTONDOWN) = ODS_BUTTONDOWN) Then
        cmdButton.SetFocus
        DoEvents   'unneeded on XP - could use If Not IsWinXPPlus() Then DoEvents
        x = (lButtonWidth - TextSize.cx + PushOffset) \ 2
        y = (lButtonHeight - TextSize.cy + PushOffset) \ 2
    Else
        x = (lButtonWidth - TextSize.cx) \ 2
        y = (lButtonHeight - TextSize.cy) \ 2
    End If
   
  'get the default WndProc address
   lpWndProC = GetProp(hWnd, PROP_LPWNDPROC)
   
  'do the default button processing
   ProcessButton = CallWindowProc(lpWndProC, hWnd, uMsg, wParam, lParam)
   
  'put our text on the button
   bRC = TextOut(lParam.hDC, x, y, sCaption, Len(sCaption))
   
  'Restore the device context to the original color
   lRC = SetTextColor(lParam.hDC, lPrevColor)
   
ProcessButton_Exit:
   Set cmdButton = Nothing

End Function


Private Sub RemoveForm(hWndParent As Long)

   Dim hWndButton As Long
   Dim cnt As Integer
   
   UnsubclassForm hWndParent
   
   On Error GoTo RemoveForm_Exit
   
   For cnt = colButtons.Count - 1 To 0 Step -1
   
      hWndButton = colButtons(cnt).hWnd
      
      If GetProp(hWndButton, PROP_HWNDPARENT) = hWndParent Then
         RemoveProp hWndButton, PROP_COLOR
         RemoveProp hWndButton, PROP_HWNDPARENT
         colButtons.Remove cnt
      End If
      
   Next cnt
   
RemoveForm_Exit:

End Sub


Private Function UnsubclassForm(hWnd As Long) As Boolean

   Dim lpWndProC As Long
   
   lpWndProC = GetProp(hWnd, PROP_LPWNDPROC)
   
   If lpWndProC = 0 Then
   
      UnsubclassForm = False
      
   Else
   
      Call SetWindowLong(hWnd, GWL_WNDPROC, lpWndProC)
      RemoveProp hWnd, PROP_LPWNDPROC
      UnsubclassForm = True
      
   End If

End Function


Private Function ButtonColorProc(ByVal hWnd As Long, _
                                 ByVal uMsg As Long, _
                                 ByVal wParam As Long, _
                                 lParam As DRAWITEMSTRUCT) As Long

   Dim lpWndProC       As Long
   Dim bProcessButton  As Boolean
   Dim sButtonKey      As String

   bProcessButton = False      'Assume default processing

   If (uMsg = WM_DRAWITEM) Then
   
     'Do we have this button? To find out, just
     'try to reference the item in the collection.
     'If it's there, we own the button.  If it's
     'not there, we'll get an error.
      sButtonKey = GetKey(lParam.hWndItem)
      bProcessButton = FindButton(sButtonKey)
   
   End If
   
   
   If bProcessButton Then
   
      ProcessButton hWnd, uMsg, wParam, lParam, sButtonKey
      
   Else
   
      lpWndProC = GetProp(hWnd, PROP_LPWNDPROC)
      ButtonColorProc = CallWindowProc(lpWndProC, hWnd, uMsg, wParam, lParam)

      If uMsg = WM_DESTROY Then RemoveForm hWnd
      
   End If

End Function


Public Function RegisterButton(Button As CommandButton, _
                               Forecolor As Long) As Boolean

   Dim hWndParent      As Long
   Dim lpWndProC       As Long
   Dim sButtonKey      As String

  'Make the colButtons key for the button
   sButtonKey = GetKey(Button.hWnd)
   
  'If we already own the button, just change the
  'color otherwise we need to process the whole thing.
   If FindButton(sButtonKey) Then
   
      SetProp Button.hWnd, PROP_COLOR, Forecolor
      Button.Refresh
      
   Else
   
     'Get the handle to the buttons parent form.
      hWndParent = GetParent(Button.hWnd)
   
     'If we can't find a parent form, report a
     'problem and get out.
      If (hWndParent = 0) Then
         RegisterButton = False
         Exit Function
      End If
   
     'found the parent, gather all of the necessary
     'button values and add it to the collection.
      colButtons.Add Button, sButtonKey
      SetProp Button.hWnd, PROP_COLOR, Forecolor
      SetProp Button.hWnd, PROP_HWNDPARENT, hWndParent
      
     'Determine if we've already subclassed this form.
      lpWndProC = GetProp(hWndParent, PROP_LPWNDPROC)
   
     'It's a new form.  Subclass it and add the
     'Window proc address to the collection.
      If (lpWndProC = 0) Then
         lpWndProC = SetWindowLong(hWndParent, _
         GWL_WNDPROC, AddressOf ButtonColorProc)
         SetProp hWndParent, PROP_LPWNDPROC, lpWndProC
      End If
   
   End If
   
   RegisterButton = True

End Function


Public Function UnregisterButton(Button As CommandButton) As Boolean

   Dim hWndParent As Long
   Dim sKeyButton As String

   sKeyButton = GetKey(Button.hWnd)

   If (FindButton(sKeyButton) = False) Then
      UnregisterButton = False
      Exit Function
   End If

   hWndParent = GetProp(Button.hWnd, PROP_HWNDPARENT)
   UnregisterButton = UnsubclassForm(hWndParent)

   colButtons.Remove sKeyButton
   RemoveProp Button.hWnd, PROP_COLOR
   RemoveProp Button.hWnd, PROP_HWNDPARENT
   
End Function


Private Function IsWinXPPlus() As Boolean

  'returns True if running WinXP (NT5.1) or later
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)

   If GetVersionEx(osv) = 1 Then
   
      IsWinXPPlus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                    (osv.dwVerMajor >= 5 And osv.dwVerMinor >= 1)

   End If

End Function




