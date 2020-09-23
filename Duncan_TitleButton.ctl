VERSION 5.00
Begin VB.UserControl Duncan_TitleButton 
   BackColor       =   &H008080FF&
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "Duncan_TitleButton.ctx":0000
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   480
      ScaleHeight     =   855
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Image picClassic 
      Height          =   210
      Index           =   3
      Left            =   4200
      Picture         =   "Duncan_TitleButton.ctx":0312
      Top             =   1560
      Width           =   240
   End
   Begin VB.Image picClassic 
      Height          =   210
      Index           =   0
      Left            =   3120
      Picture         =   "Duncan_TitleButton.ctx":05F6
      Top             =   1560
      Width           =   240
   End
   Begin VB.Image picSilver 
      Height          =   315
      Index           =   3
      Left            =   4200
      Picture         =   "Duncan_TitleButton.ctx":08DA
      Top             =   1200
      Width           =   315
   End
   Begin VB.Image picSilver 
      Height          =   315
      Index           =   2
      Left            =   3840
      Picture         =   "Duncan_TitleButton.ctx":0E5E
      Top             =   1200
      Width           =   315
   End
   Begin VB.Image picSilver 
      Height          =   315
      Index           =   1
      Left            =   3480
      Picture         =   "Duncan_TitleButton.ctx":13E2
      Top             =   1200
      Width           =   315
   End
   Begin VB.Image picSilver 
      Height          =   315
      Index           =   0
      Left            =   3120
      Picture         =   "Duncan_TitleButton.ctx":1966
      Top             =   1200
      Width           =   315
   End
   Begin VB.Image picOlive 
      Height          =   315
      Index           =   3
      Left            =   4200
      Picture         =   "Duncan_TitleButton.ctx":1EEA
      Top             =   840
      Width           =   315
   End
   Begin VB.Image picOlive 
      Height          =   315
      Index           =   2
      Left            =   3840
      Picture         =   "Duncan_TitleButton.ctx":246E
      Top             =   840
      Width           =   315
   End
   Begin VB.Image picOlive 
      Height          =   315
      Index           =   1
      Left            =   3480
      Picture         =   "Duncan_TitleButton.ctx":29F2
      Top             =   840
      Width           =   315
   End
   Begin VB.Image picOlive 
      Height          =   315
      Index           =   0
      Left            =   3120
      Picture         =   "Duncan_TitleButton.ctx":2F76
      Top             =   840
      Width           =   315
   End
   Begin VB.Image picSource 
      Height          =   315
      Index           =   3
      Left            =   4200
      Top             =   120
      Width           =   315
   End
   Begin VB.Image picSource 
      Height          =   315
      Index           =   2
      Left            =   3840
      Top             =   120
      Width           =   315
   End
   Begin VB.Image picSource 
      Height          =   315
      Index           =   1
      Left            =   3480
      Top             =   120
      Width           =   315
   End
   Begin VB.Image picSource 
      Height          =   315
      Index           =   0
      Left            =   3120
      Top             =   120
      Width           =   315
   End
   Begin VB.Image picBlue 
      Height          =   315
      Index           =   3
      Left            =   4200
      Picture         =   "Duncan_TitleButton.ctx":34FA
      Top             =   480
      Width           =   315
   End
   Begin VB.Image picBlue 
      Height          =   315
      Index           =   2
      Left            =   3840
      Picture         =   "Duncan_TitleButton.ctx":3A7E
      Top             =   480
      Width           =   315
   End
   Begin VB.Image picBlue 
      Height          =   315
      Index           =   1
      Left            =   3480
      Picture         =   "Duncan_TitleButton.ctx":4002
      Top             =   480
      Width           =   315
   End
   Begin VB.Image picBlue 
      Height          =   315
      Index           =   0
      Left            =   3120
      Picture         =   "Duncan_TitleButton.ctx":4586
      Top             =   480
      Width           =   315
   End
End
Attribute VB_Name = "Duncan_TitleButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'What does this do?
'It adds a button to the caption bar of your form.
'You only need to add one per form.

'Why?
'So you can have a "minimise to system tray" button
'designed so that everything is in one usercontrol - easy to "drop in" to new projects that way

'How?
'subclasses the picturebox, the usercontrol.
'moves the usercontrol to the parent of the form
'and then positions it so that it
'looks like it is part of the original caption bar

'Who ?
'Thanks to Paul Catton for his work on subclassing - sourced from planetsourcecode.com (54117)
'Thanks to ABSoftware for original idea, code and images - sourced from planetsourcecode.com (58679)

'When?
'Last Updated : 5 feb 2005

'Testing?
'Has only been tested on an XP machine

'To do?
'1) theme change is ok for windows themes and classic view but is no good on custom colours in classic view. should ideally change the classic button to be drawn from scratch
'2) limit to only one copy on the form at any time - cant figure out how
'3) at random times the frame of the parent can change. havent yet pinned this down

'======================================================================================================================================================
'MY DECLARES FOR THIS CONTROL
'======================================================================================================================================================
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

'so window does not appear in taskbar
Private Const WS_EX_TOOLWINDOW As Long = &H80&

'for moving the window
Private Const HWND_TOP = 0
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOSIZE = &H1


'for getting windows dimensions
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = -20
Private Const SM_CXFRAME = 32
Private Const SM_CYCAPTION = 4
Private Const SM_CXDLGFRAME = 7

'These correspond to the index of the image to be shown
Private Enum eIconStyle
    ICON_INACTIVE = 0   ' Inactive icon
    ICON_NORMAL = 1     ' Normal icon
    ICON_HOT = 2        ' When the mouse is over the icon
    ICON_MOUSEDOWN = 3  ' When the mouse is down and over the icon
End Enum

Private m_hForm As Long
Private m_Active As Boolean

Public Event Click()

'-----------------------
'DECLARATIONS FOR THEMES
'-----------------------
Private Declare Function OpenThemeData Lib "uxtheme.dll" _
   (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" _
   (ByVal hTheme As Long) As Long

Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" ( _
    ByVal pszThemeFileName As Long, _
    ByVal dwMaxNameChars As Long, _
    ByVal pszColorBuff As Long, _
    ByVal cchMaxColorChars As Long, _
    ByVal pszSizeBuff As Long, _
    ByVal cchMaxSizeChars As Long _
   ) As Long

Private Const THEME_BLUE = 1
Private Const THEME_OLIVE = 2
Private Const THEME_SILVER = 3


'======================================================================================================================================================
'SUBCLASSING DECLARES
'======================================================================================================================================================
Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum
Private Type tSubData                                                                   'Subclass data type
  hwnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
End Type
Private sc_aSubData()                As tSubData                                        'Subclass data array
Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Window Messages
Private Const WM_MOVE As Long = &H3
Private Const WM_SIZING As Long = &H214
Private Const WM_EXITSIZEMOVE As Long = &H232
Private Const WM_NCPAINT As Long = &H85
Private Const WM_SHOWWINDOW As Long = &H18
Private Const WM_ACTIVATE As Long = &H6
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_MOUSEHOVER As Long = &H2A1
Private Const WM_THEMECHANGED As Long = &H31A

'//Mouse tracking declares
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum
Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                              As Long
    dwFlags                             As TRACKMOUSEEVENT_FLAGS
    hwndTrack                           As Long
    dwHoverTime                         As Long
End Type

'Subclass handler
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
Attribute zSubclass_Proc.VB_MemberFlags = "40"
'THIS MUST BE THE FIRST PUBLIC ROUTINE IN THIS FILE.
'That includes public properties also
    Dim X As Long
    Dim Y As Long
    
    Select Case lng_hWnd
    Case m_hForm  'messages sent to the form
        Select Case uMsg
          Case WM_NCPAINT
              SetButtonPosition SWP_NOMOVE
          Case WM_MOVE, WM_SIZING
              SetButtonPosition
          Case WM_SHOWWINDOW:
              'put the button on the forms titlebar
               'lParam = 0 indicates that the message originated from a ShowWindow call
                If lParam = 0 And wParam = 0 Then 'being hidden
                   'return control of the UC to the form
                   Call SetParent(UserControl.hwnd, m_hForm)
                   Debug.Print "window parent reset to form"
                   ShowStyles
                ElseIf lParam = 0 And wParam = 1 Then 'being shown
                   'set window to have toolbar properties
                   'this prevents it showing in the taskbar
                   Call SetWindowLong(UserControl.hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW)
                   'move the control out of the form
                   Call SetParent(UserControl.hwnd, GetParent(m_hForm))
                   Debug.Print "window set to formparent"
                   ShowStyles
                   'set starting position
                   'SetButtonPosition
                End If
          Case WM_ACTIVATE
              If wParam Then  '----------------------------------- Activated
                  m_Active = True
                  SetButton ICON_NORMAL
              Else            '----------------------------------- Deactivated
                  m_Active = False
                  SetButton ICON_INACTIVE
              End If
          Case WM_THEMECHANGED
              AlignButtonsToTheme
              'change to same button in new theme
              X = picButton.Tag
              picButton.Tag = ""
              SetButton X
        End Select
    Case picButton.hwnd
        If uMsg = WM_MOUSELEAVE Then
            If m_Active Then
                SetButton ICON_NORMAL
            Else
                SetButton ICON_INACTIVE
            End If
        End If
    End Select
End Sub
'======================================================================================================================================================
'Functions
'======================================================================================================================================================
'---------
'PICBUTTON
'---------
Private Sub picButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'respond to movements
    If Button Then
        If X > picButton.ScaleLeft And _
           X < picButton.ScaleWidth And _
           Y > picButton.ScaleTop And _
           Y < picButton.ScaleHeight Then
            SetButton ICON_MOUSEDOWN
        Else
            If m_Active Then
                SetButton ICON_NORMAL
            Else
                SetButton ICON_INACTIVE
            End If
        End If
    Else
        'make sure the tool tips are in sync
        If picButton.ToolTipText <> UserControl.Extender.ToolTipText Then
            picButton.ToolTipText = UserControl.Extender.ToolTipText
        End If
        'make sure hot button is showing
        SetButton ICON_HOT
    End If
    
    Call TrackMouseLeave
End Sub
Private Sub picButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      SetButton ICON_MOUSEDOWN
   End If
   
   UserControl.Parent.SetFocus
End Sub

Private Sub picButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If X > picButton.ScaleLeft And _
           X < picButton.ScaleWidth And _
           Y > picButton.ScaleTop And _
           Y < picButton.ScaleHeight Then
            'the mouse was UP inside the control
            RaiseEvent Click
        End If
        SetButton ICON_NORMAL
    End If

End Sub

'-----------------
'PRIVATE FUNCTIONS
'-----------------
Private Function ShowStyles()

    Debug.Print "Window = Form " & m_hForm
    Debug.Print "GWL_STYLE = " & GetWindowLong(m_hForm, GWL_STYLE)
    Debug.Print "GWL_EXSTYLE = " & GetWindowLong(m_hForm, GWL_EXSTYLE)

    Debug.Print "Window = Usercontrol " & UserControl.hwnd
    Debug.Print "GWL_STYLE = " & GetWindowLong(UserControl.hwnd, GWL_STYLE)
    Debug.Print "GWL_EXSTYLE = " & GetWindowLong(UserControl.hwnd, GWL_EXSTYLE)

End Function
Private Sub SetButtonPosition(Optional lFlag As Long = SWP_FRAMECHANGED)
    'works out where on the screen the button should be placed
    'SWP_NOMOVE
    'SWP_FRAMECHANGED
    Dim R As RECT
    Dim X As Long
    Dim Y As Long
    Dim CX As Long
    Dim CY As Long
    Dim lStyle As Long
    
    'First find out where the form is
    GetWindowRect m_hForm, R
    'establish the size of the caption
    CY = GetSystemMetrics(SM_CYCAPTION)
    lStyle = GetWindowLong(m_hForm, GWL_STYLE)
    Select Case lStyle And &H80
        Case &H80:       CX = GetSystemMetrics(SM_CXDLGFRAME)
        Case Else:       CX = GetSystemMetrics(SM_CXFRAME)
    End Select
    'crop back our rectangle to exclude borders
    R.Left = R.Left + CX
    R.Right = R.Right
    R.Top = R.Top + CX
    R.Bottom = (R.Top + CY) - 1
    'R should now represent the caption bar
    'Debug.Print R.Top & "," & R.Bottom & "-" & R.Left & "," & R.Right
    'calc positioning
    X = R.Right - ((4 * (picButton.Width - 1)) + (3 * CX))
    Y = R.Top + ((R.Bottom - R.Top) - (picButton.Height - 1)) / 2
    'move window
    Call SetWindowPos(UserControl.hwnd, HWND_TOP, X, Y, picButton.Width - 1, picButton.Height - 1, lFlag)

End Sub

Private Sub SetButton(iIndex As eIconStyle)
    'changes what button is being shown
    If picButton.Tag <> CStr(iIndex) Then
        Set picButton.Picture = picSource(iIndex).Picture
        picButton.Tag = CStr(iIndex)
    End If
End Sub

Private Sub TrackMouseLeave()
    'Starts tracking the mouse
    'When the mouse leaves the control the WM_MOUSELEAVE message will be sent
    'Doesnt work for transparent windows :(
    On Error GoTo Errs
    Dim tme As TRACKMOUSEEVENT_STRUCT
    With tme
        .cbSize = Len(tme)
        .dwFlags = TME_LEAVE    'Or TME_HOVER
        .hwndTrack = picButton.hwnd
        '.dwHoverTime = HOVER_DEFAULT
    End With
    Call TrackMouseEvent(tme) '---- Track the mouse leaving the indicated window via subclassing
Errs:
End Sub

Private Function GetCurrentTheme(hwnd As Long) As Long
    'returns what Theme is currently being used by the OS
    On Error GoTo Out
    Dim hTheme As Long
    Dim sThemeFile As String, sColorName As String
    Dim lPtrThemeFile As Long
    Dim lPtrColorName As Long
    Dim hRes As Long
    
    hTheme = OpenThemeData(hwnd, StrPtr("BUTTON"))
   
    If Not (hTheme = 0) Then
        ReDim bThemeFile(0 To 260 * 2) As Byte
        lPtrThemeFile = VarPtr(bThemeFile(0))
        ReDim bColorName(0 To 260 * 2) As Byte
        lPtrColorName = VarPtr(bColorName(0))
        hRes = GetCurrentThemeName(lPtrThemeFile, 260, lPtrColorName, 260, 0, 0)
    
        sThemeFile = bThemeFile
        If InStr(LCase(sThemeFile), "luna.msstyles") Then
            sColorName = bColorName
            If InStr(LCase(sColorName), "normalcolor") Then
                GetCurrentTheme = THEME_BLUE
            End If
            If InStr(LCase(sColorName), "homestead") Then
                GetCurrentTheme = THEME_OLIVE
            End If
            If InStr(LCase(sColorName), "metallic") Then
                GetCurrentTheme = THEME_SILVER
            End If
        End If
      
        CloseThemeData hTheme
    End If
Out:

End Function

'------------
'USER CONTROL
'------------
Private Sub UserControl_Initialize()
    'position picturebox
        picButton.Left = 0
        picButton.Top = 0
    'sync buttons to theme
        AlignButtonsToTheme
    'Put a picture in the box
        SetButton ICON_NORMAL
        
End Sub
Private Sub AlignButtonsToTheme()
    'copy the appropriate button into picSource to represent the Theme being used
    Dim RGN As Long
    Dim I As Long
    Dim Theme As Long
    
    'what theme?
    Theme = GetCurrentTheme(UserControl.hwnd)
    
    'Apply buttons
    Select Case Theme
    Case THEME_BLUE
        For I = 0 To 3
            picSource(I).Picture = picBlue(I).Picture
        Next
    Case THEME_OLIVE
        For I = 0 To 3
            picSource(I).Picture = picOlive(I).Picture
        Next
    Case THEME_SILVER
        For I = 0 To 3
            picSource(I).Picture = picSilver(I).Picture
        Next
    Case Else
        For I = 0 To 2
            picSource(I).Picture = picClassic(0).Picture
        Next
        picSource(3).Picture = picClassic(3).Picture
    End Select
    
    picButton.Width = picSource(0).Width + 1
    picButton.Height = picSource(0).Height + 1
    
    'Apply region to usercontrol so that the button has correct shape
    If Theme > 0 Then
        'Initialise the picturebox to display a rounded button
        'Create the region for the round rectangle
        RGN = CreateRoundRectRgn(0, 0, picButton.Width, picButton.Width, 2, 2)
        'Apply the region
        SetWindowRgn UserControl.hwnd, RGN, True
    Else
        'Apply a blank region - will reset display area to "show all"
        'which is what we want for the square "classic view" buttons
        SetWindowRgn UserControl.hwnd, RGN, True
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim retval As Long
    
    If Ambient.UserMode Then
        'store window handles - needed later to return ownership
        m_hForm = UserControl.Parent.hwnd
        
        'Install Subclassing
        Call Subclass_Start(m_hForm)
        Call Subclass_AddMsg(m_hForm, WM_NCPAINT, MSG_AFTER)
        Call Subclass_AddMsg(m_hForm, WM_MOVE, MSG_AFTER)
        Call Subclass_AddMsg(m_hForm, WM_SIZING, MSG_AFTER)
        Call Subclass_AddMsg(m_hForm, WM_SHOWWINDOW, MSG_AFTER)
        Call Subclass_AddMsg(m_hForm, WM_ACTIVATE, MSG_AFTER)
        Call Subclass_AddMsg(m_hForm, WM_THEMECHANGED, MSG_AFTER)
        
        Call Subclass_Start(picButton.hwnd)
        Call Subclass_AddMsg(picButton.hwnd, WM_MOUSELEAVE, MSG_AFTER)
        'Call Subclass_AddMsg(picButton.hwnd, WM_MOUSEHOVER, MSG_AFTER)
    End If
End Sub

Private Sub UserControl_Resize()
    If Not Ambient.UserMode Then
        'make control tidy when designing
        UserControl.Width = picButton.Width * Screen.TwipsPerPixelX
        UserControl.Height = picButton.Height * Screen.TwipsPerPixelY
    End If
End Sub


Private Sub UserControl_InitProperties()
'Trying to limit it so that only one control can be on a form at any time
'    On Error Resume Next
'    Dim C As Control
'    For Each C In Parent.Controls
'        If TypeOf C Is Duncan_TitleButton Then
'            If C.Name = Extender.Name Then
'                Set C = Nothing
'                Debug.Print Extender.Name; " set to nothing"
'            End If
'        End If
'    Next C
End Sub

Private Sub UserControl_Terminate()
    'unload subclassing
    On Error GoTo Errs
    If Ambient.UserMode Then Call Subclass_StopAll
Errs:
End Sub


'========================================================================================
'Subclass routines below here - The programmer may call any of the following Subclass_??? routines
'======================================================================================================================================================
'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
On Error GoTo Errs

'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim I                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
'If it's the first time through here..
  If aBuf(1) = 0 Then
  
'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
           "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
           "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
           "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'Convert the string from hex pairs to bytes and store in the static machine code buffer
    I = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, I, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      I = I + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hwnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
Errs:
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
On Error GoTo Errs
  Dim I As Long
  
  I = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While I >= 0                                                                       'Iterate through each element
    With sc_aSubData(I)
      If .hwnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hwnd)                                                       'Subclass_Stop
      End If
    End With
    I = I - 1                                                                           'Next element
  Loop
Errs:
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hwnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
Errs:
End Sub

'=======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
On Error GoTo Errs
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
Errs:
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
On Error GoTo Errs
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
Errs:
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
On Error GoTo Errs
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hwnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hwnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
'  If Not bAdd Then
'    Debug.Assert False                                                                  'hWnd not found, programmer error
'  End If
Errs:

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function

'END Subclassing Code===================================================================================


