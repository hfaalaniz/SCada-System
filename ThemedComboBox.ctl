VERSION 5.00
Begin VB.UserControl ThemedComboBox 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2310
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   390
   ScaleWidth      =   2310
   Begin VB.PictureBox picImage 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      Picture         =   "ThemedComboBox.ctx":0000
      ScaleHeight     =   300
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   0
      Width           =   315
   End
   Begin VB.PictureBox picButtonState 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   1920
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   336
   End
   Begin VB.PictureBox picButtonState 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   1440
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   336
   End
   Begin VB.PictureBox picButtonState 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   960
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   312
   End
   Begin VB.PictureBox picButtonState 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   480
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   312
   End
End
Attribute VB_Name = "ThemedComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : ThemedComboBox
'    Project    : prjFrecRodamientos
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
'ThemedComboBox Control
'
'Author Ben Vonk
'10-10-2008 First version, included: Paul Caton's self Subclass v1.1.0008
'30-11-2011 Second version, Add properties so the user can customize the button

Option Explicit

' Private Constants
Private Const ALL_MESSAGES        As Long = -1
Private Const CB_GETDROPPEDSTATE  As Long = &H157
Private Const CBP_ARROWBTN        As Long = 1
Private Const GWL_WNDPROC         As Long = -4
Private Const PATCH_05            As Long = 93
Private Const PATCH_09            As Long = 137
Private Const WM_ACTIVATE         As Long = &H6
Private Const WM_COMMAND          As Long = &H111
Private Const WM_DESTROY          As Long = &H2
Private Const WM_LBUTTONDOWN      As Long = &H201
Private Const WM_LBUTTONUP        As Long = &H202
Private Const WM_MOUSEMOVE        As Long = &H200
Private Const WM_PAINT            As Long = &HF
Private Const WM_THEMECHANGED     As Long = &H31A
Private Const WM_TIMER            As Long = &H113

' Public Enumerations
Public Enum BorderColorStyles
   ThemeColors
   CustomColors
End Enum

Public Enum ButtonThemeTypes
   Windows
   User
End Enum

' Private Enumerations
Private Enum ButtonStates
   Normal
   Over
   Pressed
   Disabled
End Enum

Private Enum ControlState
   StateNormal
   StateOver
   StateFocus
   StateDown
   StateDisabled
   StateUp
End Enum

Private Enum MsgWhen
   MSG_AFTER = 1
   MSG_BEFORE = 2
   MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE
End Enum

' Private Types
Private Type OSVersionInfo
   dwOSVersionInfoSize            As Long
   dwMajorVersion                 As Long
   dwMinorVersion                 As Long
   dwBuildNumber                  As Long
   dwPlatformId                   As Long
   szCSDVersion                   As String * 128
End Type

Private Type PointAPI
   X                              As Long
   Y                              As Long
End Type

Private Type Rect
   Left                           As Long
   Top                            As Long
   Right                          As Long
   Bottom                         As Long
End Type

Private Type ComboBoxInfo
   cbSize                         As Long
   rcItem                         As Rect
   rcButton                       As Rect
   lStateButton                   As Long
   hWndCombo                      As Long
   hWndEdit                       As Long
   hWndList                       As Long
End Type

Private Type SubclassDataType
   hWnd                           As Long
   nAddrSclass                    As Long
   nAddrOrig                      As Long
   nMsgCountA                     As Long
   nMsgCountB                     As Long
   aMsgTabelA()                   As Long
   aMsgTabelB()                   As Long
End Type

' Private Variables
Private ButtonDown                As Boolean
Private IsThemed                  As Boolean
Private IsThemedWindows           As Boolean
Private m_Activated               As Boolean
Private MouseOver                 As Boolean
Private m_BorderColorStyle        As BorderColorStyles
Private m_ButtonThemeType         As ButtonThemeTypes
Private ButtonState               As ControlState
Private DefaultBorderColor        As Long
Private m_ComboBoxBorderColor     As Long
Private m_DriveListBoxBorderColor As Long
Private m_ImageComboBorderColor   As Long
Private SubclassMemory            As Long
Private TimerID                   As Long
Private SubclassData()            As SubclassDataType

' Private API's
Private Declare Function CreatePen Lib "GDI32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LineTo Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As PointAPI) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVersionInfo) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Private Declare Function FrameRect Lib "User32" (ByVal hDC As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Private Declare Function GetComboBoxInfo Lib "User32" (ByVal hWndCombo As Long, ByRef pcbi As ComboBoxInfo) As Long
Private Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function GetSysColor Lib "User32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function IsWindowEnabled Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function KillTimer Lib "User32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function MoveWindow Lib "User32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetTimer Lib "User32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function SetWindowLongA Lib "User32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function CloseThemeData Lib "UxTheme" (ByVal lngTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "UxTheme" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Rect, pClipRect As Rect) As Long
Private Declare Function GetCurrentThemeName Lib "UxTheme" (ByVal pszThemeFileName As Long, ByVal cchMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function GetThemeDocumentationProperty Lib "UxTheme" (ByVal pszThemeName As Long, ByVal pszPropertyName As Long, ByVal pszValueBuff As Long, ByVal cchMaxValChars As Long) As Long
Private Declare Function OpenThemeData Lib "UxTheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub Subclass_WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lhWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    '<EhHeader>
    On Error GoTo Subclass_WndProc_Err
    '</EhHeader>

Const CBN_CLOSEUP  As Long = 8

Static lngHandle   As Long

Dim lngListWindow  As Long

   Select Case uMsg
      Case WM_ACTIVATE
         If Not m_Activated Then If IsThemed Then Call Initialize
         
      Case WM_COMMAND
         If lngHandle = lParam Then
            If wParam \ &H10000 = CBN_CLOSEUP Then
               If ButtonDown Then ButtonDown = False
               
               ButtonState = StateNormal
               MouseOver = False
               TimerID = KillTimer(lhWnd, TimerID)
            End If
            
            Call DrawComboBox(lngHandle)
         End If
         
      Case WM_DESTROY
         Call Subclass_Stop(lhWnd)
         
      Case WM_LBUTTONDOWN
         If lhWnd = lngHandle Then
            ButtonState = StateDown
            
            Call DrawComboBox(lhWnd)
         End If
         
      Case WM_LBUTTONUP
         If lhWnd = lngHandle Then
            ButtonState = StateUp
            
            Call DrawComboBox(lhWnd)
         End If
         
      Case WM_MOUSEMOVE
         If InControl(lhWnd) Then
            lngHandle = lhWnd
            
            If Not MouseOver Then
               MouseOver = True
               ButtonState = StateOver
               TimerID = SetTimer(lhWnd, TimerID, 1, SubclassData(Subclass_Index(lhWnd)).nAddrSclass)
               
               Call DrawComboBox(lhWnd)
            End If
            
         Else
            ButtonState = StateDown
         End If
         
      Case WM_PAINT
         GetComboBoxButton lhWnd, lngListWindow
         
         If lhWnd = lngListWindow Then
            Call DrawComboBoxListWindow(lhWnd)
            
         Else
            Call DrawComboBox(lhWnd)
         End If
         
      Case WM_THEMECHANGED
         IsThemed = CheckIsThemed
         
      Case WM_TIMER
         If InControl(lhWnd) Then
            MouseOver = True
            
            If (ButtonState <> StateDown) And SendMessage(lhWnd, CB_GETDROPPEDSTATE, 0, ByVal 0&) Then ButtonState = StateOver
            
         Else
            MouseOver = False
            ButtonState = StateNormal
            TimerID = KillTimer(lhWnd, TimerID)
            
            Call DrawComboBox(lhWnd)
         End If
   End Select

    '<EhFooter>
    Exit Sub

Subclass_WndProc_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.Subclass_WndProc", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub

Private Function Subclass_AddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    '<EhHeader>
    On Error GoTo Subclass_AddrFunc_Err
    '</EhHeader>

   Subclass_AddrFunc = GetProcAddress(GetModuleHandle(sDLL), sProc)
   Debug.Assert Subclass_AddrFunc

    '<EhFooter>
    Exit Function

Subclass_AddrFunc_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.Subclass_AddrFunc", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Function

Private Function Subclass_Index(ByVal lhWnd As Long, Optional ByVal bAdd As Boolean) As Long
    '<EhHeader>
    On Error GoTo Subclass_Index_Err
    '</EhHeader>

   For Subclass_Index = UBound(SubclassData) To 0 Step -1
      If SubclassData(Subclass_Index).hWnd = lhWnd Then
         If Not bAdd Then Exit Function
         
      ElseIf SubclassData(Subclass_Index).hWnd = 0 Then
         If bAdd Then Exit Function
      End If
   Next 'Subclass_Index
   
   If Not bAdd Then Debug.Assert False

    '<EhFooter>
    Exit Function

Subclass_Index_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.Subclass_Index", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Function

Private Function Subclass_InIDE() As Boolean
    '<EhHeader>
    On Error GoTo Subclass_InIDE_Err
    '</EhHeader>

   Debug.Assert Subclass_SetTrue(Subclass_InIDE)

    '<EhFooter>
    Exit Function

Subclass_InIDE_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.Subclass_InIDE", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Function

Private Function Subclass_Initialize(ByVal lhWnd As Long) As Long
    '<EhHeader>
    On Error GoTo Subclass_Initialize_Err
    '</EhHeader>

Const CODE_LEN                  As Long = 200
Const GMEM_FIXED                As Long = 0
Const PATCH_01                  As Long = 18
Const PATCH_02                  As Long = 68
Const PATCH_03                  As Long = 78
Const PATCH_06                  As Long = 116
Const PATCH_07                  As Long = 121
Const PATCH_0A                  As Long = 186
Const FUNC_CWP                  As String = "CallWindowProcA"
Const FUNC_EBM                  As String = "EbMode"
Const FUNC_SWL                  As String = "SetWindowLongA"
Const MOD_USER                  As String = "User32"
Const MOD_VBA5                  As String = "vba5"
Const MOD_VBA6                  As String = "vba6"

Static bytBuffer(1 To CODE_LEN) As Byte
Static lngCWP                   As Long
Static lngEbMode                As Long
Static lngSWL                   As Long

Dim lngCount                    As Long
Dim lngIndex                    As Long
Dim strHex                      As String

   If bytBuffer(1) Then
      lngIndex = Subclass_Index(lhWnd, True)
      
      If lngIndex = -1 Then
         lngIndex = UBound(SubclassData) + 1
         
         ReDim Preserve SubclassData(lngIndex) As SubclassDataType
      End If
      
      Subclass_Initialize = lngIndex
      
   Else
      strHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
      
      For lngCount = 1 To CODE_LEN
         bytBuffer(lngCount) = Val("&H" & Left(strHex, 2))
         strHex = Mid(strHex, 3)
      Next 'lngCount
      
      If Subclass_InIDE Then
         bytBuffer(16) = &H90
         bytBuffer(17) = &H90
         lngEbMode = Subclass_AddrFunc(MOD_VBA6, FUNC_EBM)
         
         If lngEbMode = 0 Then lngEbMode = Subclass_AddrFunc(MOD_VBA5, FUNC_EBM)
      End If
      
      lngCWP = Subclass_AddrFunc(MOD_USER, FUNC_CWP)
      lngSWL = Subclass_AddrFunc(MOD_USER, FUNC_SWL)
      
      ReDim SubclassData(0) As SubclassDataType
   End If
   
   With SubclassData(lngIndex)
      .hWnd = lhWnd
      .nAddrSclass = GlobalAlloc(GMEM_FIXED, CODE_LEN)
      .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSclass)
      
      Call CopyMemory(ByVal .nAddrSclass, bytBuffer(1), CODE_LEN)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_01, lngEbMode)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_02, .nAddrOrig)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_03, lngSWL)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_06, .nAddrOrig)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_07, lngCWP)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_0A, ObjPtr(Me))
   End With

    '<EhFooter>
    Exit Function

Subclass_Initialize_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.Subclass_Initialize", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Function

Private Function Subclass_SetTrue(ByRef bValue As Boolean) As Boolean
    '<EhHeader>
    On Error GoTo Subclass_SetTrue_Err
    '</EhHeader>

   Subclass_SetTrue = True
   bValue = True

    '<EhFooter>
    Exit Function

Subclass_SetTrue_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.Subclass_SetTrue", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Function

Private Sub Subclass_AddMsg(ByVal lhWnd As Long, ByVal uMsg As Long, Optional ByVal When As MsgWhen = MSG_AFTER)
    '<EhHeader>
    On Error GoTo Subclass_AddMsg_Err
    '</EhHeader>

   With SubclassData(Subclass_Index(lhWnd))
      If When And MSG_BEFORE Then Call Subclass_DoAddMsg(uMsg, .aMsgTabelB, .nMsgCountB, MSG_BEFORE, .nAddrSclass)
      If When And MSG_AFTER Then Call Subclass_DoAddMsg(uMsg, .aMsgTabelA, .nMsgCountA, MSG_AFTER, .nAddrSclass)
   End With

    '<EhFooter>
    Exit Sub

Subclass_AddMsg_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.Subclass_AddMsg", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub

Private Sub Subclass_DoAddMsg(ByVal uMsg As Long, ByRef aMsgTabel() As Long, ByRef nMsgCount As Long, ByVal When As MsgWhen, ByVal nAddr As Long)
    '<EhHeader>
    On Error GoTo Subclass_DoAddMsg_Err
    '</EhHeader>

Const PATCH_04 As Long = 88
Const PATCH_08 As Long = 132

Dim lngEntry   As Long

   ReDim lngOffset(1) As Long
   
   If uMsg = ALL_MESSAGES Then
      nMsgCount = ALL_MESSAGES
      
   Else
      For lngEntry = 1 To nMsgCount - 1
         If aMsgTabel(lngEntry) = 0 Then
            aMsgTabel(lngEntry) = uMsg
            
            GoTo ExitSub
            
         ElseIf aMsgTabel(lngEntry) = uMsg Then
            GoTo ExitSub
         End If
      Next 'lngEntry
      
      nMsgCount = nMsgCount + 1
      
      ReDim Preserve aMsgTabel(1 To nMsgCount) As Long
      
      aMsgTabel(nMsgCount) = uMsg
   End If
   
   If When = MSG_BEFORE Then
      lngOffset(0) = PATCH_04
      lngOffset(1) = PATCH_05
      
   Else
      lngOffset(0) = PATCH_08
      lngOffset(1) = PATCH_09
   End If
   
   If uMsg <> ALL_MESSAGES Then Call Subclass_PatchVal(nAddr, lngOffset(0), VarPtr(aMsgTabel(1)))
   
   Call Subclass_PatchVal(nAddr, lngOffset(1), nMsgCount)
   
ExitSub:
   Erase lngOffset

    '<EhFooter>
    Exit Sub

Subclass_DoAddMsg_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.Subclass_DoAddMsg", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub

Private Sub Subclass_PatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    '<EhHeader>
    On Error GoTo Subclass_PatchRel_Err
    '</EhHeader>

   Call CopyMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)

    '<EhFooter>
    Exit Sub

Subclass_PatchRel_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.Subclass_PatchRel", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub

Private Sub Subclass_PatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    '<EhHeader>
    On Error GoTo Subclass_PatchVal_Err
    '</EhHeader>

   Call CopyMemory(ByVal nAddr + nOffset, nValue, 4)

    '<EhFooter>
    Exit Sub

Subclass_PatchVal_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.Subclass_PatchVal", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub

Private Sub Subclass_Stop(ByVal lhWnd As Long)
    '<EhHeader>
    On Error GoTo Subclass_Stop_Err
    '</EhHeader>

   With SubclassData(Subclass_Index(lhWnd))
      SetWindowLongA .hWnd, GWL_WNDPROC, .nAddrOrig
      
      Call Subclass_PatchVal(.nAddrSclass, PATCH_05, 0)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_09, 0)
      
      GlobalFree .nAddrSclass
      .hWnd = 0
      .nMsgCountA = 0
      .nMsgCountB = 0
      Erase .aMsgTabelA, .aMsgTabelB
   End With

    '<EhFooter>
    Exit Sub

Subclass_Stop_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.Subclass_Stop", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub

Private Sub Subclass_Terminate()
    '<EhHeader>
    On Error GoTo Subclass_Terminate_Err
    '</EhHeader>

Dim lngCount As Long

   For lngCount = UBound(SubclassData) To 0 Step -1
      If SubclassData(lngCount).hWnd Then Call Subclass_Stop(SubclassData(lngCount).hWnd)
   Next 'lngCount

    '<EhFooter>
    Exit Sub

Subclass_Terminate_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.Subclass_Terminate", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub

Public Property Get BorderColorStyle() As BorderColorStyles
Attribute BorderColorStyle.VB_Description = "Returns/sets the border style for an object."
    '<EhHeader>
    On Error GoTo BorderColorStyle_Err
    '</EhHeader>

   BorderColorStyle = m_BorderColorStyle

    '<EhFooter>
    Exit Property

BorderColorStyle_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.BorderColorStyle", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Let BorderColorStyle(ByVal NewBorderColorStyle As BorderColorStyles)
    '<EhHeader>
    On Error GoTo BorderColorStyle_Err
    '</EhHeader>

   If NewBorderColorStyle < ThemeColors Then NewBorderColorStyle = ThemeColors
   If NewBorderColorStyle > CustomColors Then NewBorderColorStyle = CustomColors
   
   m_BorderColorStyle = NewBorderColorStyle
   PropertyChanged "BorderColorStyle"

    '<EhFooter>
    Exit Property

BorderColorStyle_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.BorderColorStyle", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Get ButtonDisabled() As StdPicture
Attribute ButtonDisabled.VB_Description = "Returns/sets a graphic to be displayed when the control is disabled. (Only if ButtonThemeType is set to User!)"
    '<EhHeader>
    On Error GoTo ButtonDisabled_Err
    '</EhHeader>

   Set ButtonDisabled = picButtonState.item(Disabled).Picture

    '<EhFooter>
    Exit Property

ButtonDisabled_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ButtonDisabled", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Let ButtonDisabled(ByVal NewButtonDisabled As StdPicture)
    '<EhHeader>
    On Error GoTo ButtonDisabled_Err
    '</EhHeader>

   Set ButtonDisabled = NewButtonDisabled

    '<EhFooter>
    Exit Property

ButtonDisabled_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ButtonDisabled", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Set ButtonDisabled(ByVal NewButtonDisabled As StdPicture)
    '<EhHeader>
    On Error GoTo ButtonDisabled_Err
    '</EhHeader>

   picButtonState.item(Disabled).Picture = NewButtonDisabled
   PropertyChanged "ButtonDisabled"

    '<EhFooter>
    Exit Property

ButtonDisabled_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ButtonDisabled", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Get ButtonNormal() As StdPicture
Attribute ButtonNormal.VB_Description = "Returns/sets a graphic to be displayed in an button normal state of the control. (Only if ButtonThemeType is set to User!)"
    '<EhHeader>
    On Error GoTo ButtonNormal_Err
    '</EhHeader>

   Set ButtonNormal = picButtonState.item(Normal).Picture

    '<EhFooter>
    Exit Property

ButtonNormal_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ButtonNormal", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Let ButtonNormal(ByVal NewButtonNormal As StdPicture)
    '<EhHeader>
    On Error GoTo ButtonNormal_Err
    '</EhHeader>

   Set ButtonNormal = NewButtonNormal

    '<EhFooter>
    Exit Property

ButtonNormal_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ButtonNormal", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Set ButtonNormal(ByVal NewButtonNormal As StdPicture)
    '<EhHeader>
    On Error GoTo ButtonNormal_Err
    '</EhHeader>

   picButtonState.item(Normal).Picture = NewButtonNormal
   PropertyChanged "ButtonNormal"
   
   Call CheckButtonThemeType

    '<EhFooter>
    Exit Property

ButtonNormal_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ButtonNormal", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Get ButtonOver() As StdPicture
Attribute ButtonOver.VB_Description = "Returns/sets a graphic to be displayed in an button over state of the control. (Only if ButtonThemeType is set to User!)"
    '<EhHeader>
    On Error GoTo ButtonOver_Err
    '</EhHeader>

   Set ButtonOver = picButtonState.item(Over).Picture

    '<EhFooter>
    Exit Property

ButtonOver_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ButtonOver", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Let ButtonOver(ByVal NewButtonOver As StdPicture)
    '<EhHeader>
    On Error GoTo ButtonOver_Err
    '</EhHeader>

   Set ButtonOver = NewButtonOver

    '<EhFooter>
    Exit Property

ButtonOver_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ButtonOver", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Set ButtonOver(ByVal NewButtonOver As StdPicture)
    '<EhHeader>
    On Error GoTo ButtonOver_Err
    '</EhHeader>

   picButtonState.item(Over).Picture = NewButtonOver
   PropertyChanged "ButtonOver"

    '<EhFooter>
    Exit Property

ButtonOver_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ButtonOver", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Get ButtonPressed() As StdPicture
Attribute ButtonPressed.VB_Description = "Returns/sets a graphic to be displayed in an button pressed state of the control. (Only if ButtonThemeType is set to User!)"
    '<EhHeader>
    On Error GoTo ButtonPressed_Err
    '</EhHeader>

   Set ButtonPressed = picButtonState.item(Pressed).Picture

    '<EhFooter>
    Exit Property

ButtonPressed_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ButtonPressed", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Let ButtonPressed(ByVal NewButtonPressed As StdPicture)
    '<EhHeader>
    On Error GoTo ButtonPressed_Err
    '</EhHeader>

   Set ButtonPressed = NewButtonPressed

    '<EhFooter>
    Exit Property

ButtonPressed_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ButtonPressed", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Set ButtonPressed(ByVal NewButtonPressed As StdPicture)
    '<EhHeader>
    On Error GoTo ButtonPressed_Err
    '</EhHeader>

   picButtonState.item(Pressed).Picture = NewButtonPressed
   PropertyChanged "ButtonPressed"

    '<EhFooter>
    Exit Property

ButtonPressed_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ButtonPressed", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Get ButtonThemeType() As ButtonThemeTypes
Attribute ButtonThemeType.VB_Description = "Returns/sets a theme type of the ThemedComboBox control."
    '<EhHeader>
    On Error GoTo ButtonThemeType_Err
    '</EhHeader>

   ButtonThemeType = m_ButtonThemeType

    '<EhFooter>
    Exit Property

ButtonThemeType_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ButtonThemeType", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Let ButtonThemeType(ByVal NewButtonThemeType As ButtonThemeTypes)
    '<EhHeader>
    On Error GoTo ButtonThemeType_Err
    '</EhHeader>

   m_ButtonThemeType = NewButtonThemeType
   PropertyChanged "ButtonThemeType"
   
   Call CheckButtonThemeType

    '<EhFooter>
    Exit Property

ButtonThemeType_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ButtonThemeType", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Get ComboBoxBorderColor() As OLE_COLOR
Attribute ComboBoxBorderColor.VB_Description = "Returns/sets the color of an ComboBox border."
    '<EhHeader>
    On Error GoTo ComboBoxBorderColor_Err
    '</EhHeader>

   ComboBoxBorderColor = m_ComboBoxBorderColor

    '<EhFooter>
    Exit Property

ComboBoxBorderColor_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ComboBoxBorderColor", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Let ComboBoxBorderColor(ByVal NewComboBoxBorderColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo ComboBoxBorderColor_Err
    '</EhHeader>

   m_ComboBoxBorderColor = NewComboBoxBorderColor
   PropertyChanged "ComboBoxBorderColor"

    '<EhFooter>
    Exit Property

ComboBoxBorderColor_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ComboBoxBorderColor", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Get DriveListBoxBorderColor() As OLE_COLOR
Attribute DriveListBoxBorderColor.VB_Description = "Returns/sets the color of an DriveListBox border."
    '<EhHeader>
    On Error GoTo DriveListBoxBorderColor_Err
    '</EhHeader>

   DriveListBoxBorderColor = m_DriveListBoxBorderColor

    '<EhFooter>
    Exit Property

DriveListBoxBorderColor_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.DriveListBoxBorderColor", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Let DriveListBoxBorderColor(ByVal NewDriveListBoxBorderColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo DriveListBoxBorderColor_Err
    '</EhHeader>

   m_DriveListBoxBorderColor = NewDriveListBoxBorderColor
   PropertyChanged "DriveListBoxBorderColor"

    '<EhFooter>
    Exit Property

DriveListBoxBorderColor_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.DriveListBoxBorderColor", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Get ImageComboBorderColor() As OLE_COLOR
Attribute ImageComboBorderColor.VB_Description = "Returns/sets the color of an ImageCombo border."
    '<EhHeader>
    On Error GoTo ImageComboBorderColor_Err
    '</EhHeader>

   ImageComboBorderColor = m_ImageComboBorderColor

    '<EhFooter>
    Exit Property

ImageComboBorderColor_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ImageComboBorderColor", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Property Let ImageComboBorderColor(ByVal NewImageComboBoxBorderColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo ImageComboBorderColor_Err
    '</EhHeader>

   m_ImageComboBorderColor = NewImageComboBoxBorderColor
   PropertyChanged "ImageComboBorderColor"

    '<EhFooter>
    Exit Property

ImageComboBorderColor_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.ImageComboBorderColor", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Property

Public Function Activated() As Boolean
    '<EhHeader>
    On Error GoTo Activated_Err
    '</EhHeader>

   Activated = m_Activated

    '<EhFooter>
    Exit Function

Activated_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.Activated", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Function

Private Function CheckIsComboBox(ByRef hWnd As Long, Optional ByRef ComboBoxBorderColor As Long) As Boolean
    '<EhHeader>
    On Error GoTo CheckIsComboBox_Err
    '</EhHeader>

Dim strClassName As String * 255

   Select Case Left(strClassName, GetClassName(hWnd, strClassName, Len(strClassName)))
      Case "ImageCombo20WndClass"
         CheckIsComboBox = True
         ComboBoxBorderColor = m_ImageComboBorderColor
         hWnd = FindWindowEx(hWnd, 0, "ComboBox", ByVal 0&)
         
      Case "ThunderComboBox", "ThunderRT6ComboBox"
         CheckIsComboBox = True
         ComboBoxBorderColor = m_ComboBoxBorderColor
         
      Case "ThunderDriveListBox", "ThunderRT6DriveListBox"
         CheckIsComboBox = True
         ComboBoxBorderColor = m_DriveListBoxBorderColor
   End Select

    '<EhFooter>
    Exit Function

CheckIsComboBox_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.CheckIsComboBox", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Function

Private Function CheckIsThemed() As Boolean
    '<EhHeader>
    On Error GoTo CheckIsThemed_Err
    '</EhHeader>

Const VER_PLATFORM_WIN32_NT As Long = 2

Dim lngLibrary              As Long
Dim osvInfo                 As OSVersionInfo
Dim strName                 As String
Dim strTheme                As String

   IsThemedWindows = False
   
   With osvInfo
      .dwOSVersionInfoSize = Len(osvInfo)
      GetVersionEx osvInfo
      
      If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
         If ((.dwMajorVersion > 4) And .dwMinorVersion) Or (.dwMajorVersion > 5) Then
            IsThemedWindows = True
            lngLibrary = LoadLibrary("UxTheme")
            
            If lngLibrary Then
               strTheme = String(255, vbNullChar)
               GetCurrentThemeName StrPtr(strTheme), Len(strTheme), 0, 0, 0, 0
               strTheme = StripNull(strTheme)
               
               If Len(strTheme) Then
                  strName = String(255, vbNullChar)
                  GetThemeDocumentationProperty StrPtr(strTheme), StrPtr("ThemeName"), StrPtr(strName), Len(strName)
                  CheckIsThemed = (StripNull(strName) <> "")
               End If
               
               FreeLibrary lngLibrary
            End If
         End If
      End If
   End With

    '<EhFooter>
    Exit Function

CheckIsThemed_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.CheckIsThemed", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Function

Private Function GetComboBoxButton(ByVal hWnd As Long, Optional ByRef ListWindow As Long, Optional ByRef ButtonWidth As Long) As Boolean
    '<EhHeader>
    On Error GoTo GetComboBoxButton_Err
    '</EhHeader>

Dim cbiCombo As ComboBoxInfo

   With cbiCombo
      .cbSize = Len(cbiCombo)
      GetComboBoxInfo hWnd, cbiCombo
      ListWindow = .hWndList
      ButtonWidth = .rcButton.Right - .rcButton.Left + 1
      GetComboBoxButton = (.lStateButton <> &H8000&)
   End With

    '<EhFooter>
    Exit Function

GetComboBoxButton_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.GetComboBoxButton", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Function

Private Function GetDefaultBorderColor() As Long
    '<EhHeader>
    On Error GoTo GetDefaultBorderColor_Err
    '</EhHeader>

Const EDP_EDITTEXT As Long = 1
Const EDS_ASSIST   As Long = 1

Dim lngTheme       As Long
Dim rctWindow      As Rect

   If IsThemedWindows Then
      rctWindow.Right = 4
      rctWindow.Bottom = 4
      lngTheme = OpenThemeData(hWnd, StrPtr("Edit"))
      DrawThemeBackground lngTheme, hDC, EDP_EDITTEXT, EDS_ASSIST, rctWindow, rctWindow
      CloseThemeData lngTheme
   End If
   
   GetDefaultBorderColor = GetPixel(hDC, 0, 0)

    '<EhFooter>
    Exit Function

GetDefaultBorderColor_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.GetDefaultBorderColor", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Function

Private Function GetLongColor(ByVal Color As Long) As Long
    '<EhHeader>
    On Error GoTo GetLongColor_Err
    '</EhHeader>

   If Color And &H80000000 Then
      GetLongColor = GetSysColor(Color And &H7FFFFFFF)
      
   Else
      GetLongColor = Color
   End If

    '<EhFooter>
    Exit Function

GetLongColor_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.GetLongColor", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Function

Private Function InControl(ByVal hWnd As Long) As Boolean
    '<EhHeader>
    On Error GoTo InControl_Err
    '</EhHeader>

Dim ptaMouse As PointAPI

   GetCursorPos ptaMouse
   InControl = (WindowFromPoint(ptaMouse.X, ptaMouse.Y) = hWnd)

    '<EhFooter>
    Exit Function

InControl_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.InControl", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Function

Private Function StripNull(ByVal Text As String) As String
    '<EhHeader>
    On Error GoTo StripNull_Err
    '</EhHeader>

   StripNull = Left(Text, StrLen(StrPtr(Text)))

    '<EhFooter>
    Exit Function

StripNull_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.StripNull", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Function

Private Sub CheckButtonThemeType()
    '<EhHeader>
    On Error GoTo CheckButtonThemeType_Err
    '</EhHeader>

   If picButtonState(Normal).Picture = 0 And (m_ButtonThemeType = User) Then ButtonThemeType = Windows

    '<EhFooter>
    Exit Sub

CheckButtonThemeType_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.CheckButtonThemeType", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub

Private Sub DrawBorder(ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal Color As Long)
    '<EhHeader>
    On Error GoTo DrawBorder_Err
    '</EhHeader>

Dim lngBrush As Long
Dim rctFrame As Rect

   With rctFrame
      .Top = Top
      .Left = Left
      .Right = Left + Right
      .Bottom = Top + Bottom
   End With
   
   ' Draw the border around the control with the given color
   lngBrush = CreateSolidBrush(Color)
   FrameRect hDC, rctFrame, lngBrush
   DeleteObject lngBrush

    '<EhFooter>
    Exit Sub

DrawBorder_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.DrawBorder", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub

Private Sub DrawComboBox(ByVal hWnd As Long)
    '<EhHeader>
    On Error GoTo DrawComboBox_Err
    '</EhHeader>

Const CBXS_DISABLED As Long = 4
Const CBXS_HOT      As Long = 2
Const CBXS_NORMAL   As Long = 1
Const CBXS_PRESSED  As Long = 3

Dim blnHasButton    As Boolean
Dim intBorderLine   As Integer
Dim intLine         As Integer
Dim lngButtonWidth  As Long
Dim lngColor(1)     As Long
Dim lngDC           As Long
Dim lngStateID      As Long
Dim lngTheme        As Long
Dim lngWindow       As Long
Dim rctClient       As Rect

   ' StateDisabled
   If IsWindowEnabled(hWnd) = 0 Then
      lngStateID = CBXS_DISABLED
      
   ElseIf ButtonState = StateOver Then
      If ButtonDown Then
         lngStateID = CBXS_PRESSED
         
      Else
         lngStateID = CBXS_HOT
      End If
      
   ElseIf ButtonState = StateDown Then
      lngStateID = CBXS_PRESSED
      ButtonDown = True
      
   ElseIf ButtonState = StateUp Then
      If InControl(hWnd) Then
         lngStateID = CBXS_HOT
         
      Else
         lngStateID = CBXS_NORMAL
      End If
      
      ButtonDown = False
      
   ' StateNormal or StateFocus
   ElseIf ButtonDown Then
      lngStateID = CBXS_PRESSED
      
   Else
      lngStateID = CBXS_NORMAL
   End If
   
   If Not ButtonDown And SendMessage(hWnd, CB_GETDROPPEDSTATE, 0, ByVal 0&) Then lngStateID = CBXS_NORMAL
   
   lngDC = GetDC(hWnd)
   blnHasButton = GetComboBoxButton(hWnd, , lngButtonWidth)
   GetClientRect hWnd, rctClient
   lngColor(1) = GetPixel(lngDC, 2, 2)
   lngWindow = FindWindowEx(hWnd, 0, "Edit", ByVal 0&)
   
   If m_BorderColorStyle = ThemeColors Then
      lngColor(0) = DefaultBorderColor
      
   Else
      CheckIsComboBox hWnd, lngColor(0)
   End If
   
   With rctClient
      For intLine = 0 To 1
         Call DrawLine(lngDC, .Right - lngButtonWidth - intLine, 2, .Right - lngButtonWidth - intLine, .Bottom - 2, lngColor(1))
      Next 'intLine
      
      If Not blnHasButton Then
         intBorderLine = 21 + (3 And (Screen.TwipsPerPixelY = 12))
         
         For intLine = 19 To 25
            Call DrawLine(lngDC, 0, .Top + intLine, .Right, .Top + intLine, lngColor(1 - (1 And (intLine = intBorderLine))))
         Next 'intLine
         
      ElseIf lngWindow Then
         MoveWindow lngWindow, .Left + 3, .Top + 3, .Right - lngButtonWidth - 3, .Bottom - 5, 0
      End If
      
      Call DrawBorder(lngDC, 1, 1, .Right - 2, .Bottom - 2, lngColor(1))
      Call DrawBorder(lngDC, 0, 0, .Right, .Bottom, lngColor(0))
      
      If blnHasButton Then
         .Top = 1
         .Left = .Right - lngButtonWidth
         .Right = .Right - 1
         .Bottom = .Bottom - 1
         
         If m_ButtonThemeType = Windows Then
            lngTheme = OpenThemeData(hWnd, StrPtr("ComboBox"))
            DrawThemeBackground lngTheme, lngDC, CBP_ARROWBTN, lngStateID, rctClient, rctClient
            CloseThemeData lngTheme
            
         Else
            'lngStateID = 0 - StateNormal
            'lngStateID = 1 - StateOver
            'lngStateID = 2 - StateDown
            'lngStateID = 3 - StateDisabled
            lngStateID = lngStateID - 1
            StretchBlt lngDC, .Left, .Top, .Right - .Left, .Bottom - .Top, picButtonState.item(lngStateID).hDC, 0, 0, picButtonState.item(lngStateID).ScaleWidth, picButtonState.item(lngStateID).ScaleHeight, vbSrcCopy
         End If
      End If
   End With
   
   DeleteDC hWnd
   Erase lngColor

    '<EhFooter>
    Exit Sub

DrawComboBox_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.DrawComboBox", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub

Private Sub DrawComboBoxListWindow(ByVal hWnd As Long)
    '<EhHeader>
    On Error GoTo DrawComboBoxListWindow_Err
    '</EhHeader>

Const GWL_EXSTYLE      As Long = -20
Const GWL_STYLE        As Long = -16
Const SWP_FRAMECHANGED As Long = &H20
Const SWP_NOACTIVATE   As Long = &H10
Const SWP_NOMOVE       As Long = &H2
Const SWP_NOSIZE       As Long = &H1
Const SWP_NOZORDER     As Long = &H4
Const WS_BORDER        As Long = &H800000
Const WS_EX_CLIENTEDGE As Long = &H200

Dim lngParent          As Long
Dim lngTop             As Long
Dim rctClient(1)       As Rect

   lngParent = GetParent(hWnd)
   GetClientRect lngParent, rctClient(0)
   GetClientRect hWnd, rctClient(1)
   
   With rctClient(1)
      ' Move the ComboBox ListWindow
      lngTop = rctClient(0).Bottom - .Bottom - 2
      MoveWindow hWnd, .Left + 1, lngTop, rctClient(0).Right - 2, .Bottom + lngTop - 7, 0
   End With
   
   ' Make the conrol flat
   SetWindowLongA hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) And Not WS_BORDER
   SetWindowLongA hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_CLIENTEDGE
   SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
   Erase rctClient
   
   Call DrawComboBox(hWnd)
   ' No more subclassing needed for this item
   Call Subclass_Stop(hWnd)

    '<EhFooter>
    Exit Sub

DrawComboBoxListWindow_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.DrawComboBoxListWindow", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub

Public Sub DrawLine(ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
    '<EhHeader>
    On Error GoTo DrawLine_Err
    '</EhHeader>

Dim lngPen(1) As Long
Dim ptaTemp   As PointAPI

   ' Draw a line in the control with the given color
   lngPen(0) = CreatePen(0, 1, GetLongColor(Color))
   lngPen(1) = SelectObject(hDC, lngPen(0))
   MoveToEx hDC, X1, Y1, ptaTemp
   LineTo hDC, X2, Y2
   SelectObject hDC, lngPen(1)
   DeleteObject lngPen(1)
   DeleteObject lngPen(0)
   Erase lngPen

    '<EhFooter>
    Exit Sub

DrawLine_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.DrawLine", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub

Private Sub Initialize()
    '<EhHeader>
    On Error GoTo Initialize_Err
    '</EhHeader>

Dim ctlControl As control
Dim lngWindow  As Long

   If Ambient.UserMode Then
      On Local Error Resume Next
      
      ' Search for all ComboBoxes on the Parent
      For Each ctlControl In Parent.controls
         Err.Clear
         m_Activated = True
         lngWindow = ctlControl.hWnd
         
         If CheckIsComboBox(lngWindow) Then
            Call Subclass_Initialize(lngWindow)
            Call Subclass_AddMsg(lngWindow, WM_COMMAND)
            Call Subclass_AddMsg(lngWindow, WM_DESTROY, MSG_BEFORE)
            Call Subclass_AddMsg(lngWindow, WM_LBUTTONDOWN, MSG_BEFORE)
            Call Subclass_AddMsg(lngWindow, WM_LBUTTONUP)
            Call Subclass_AddMsg(lngWindow, WM_MOUSEMOVE)
            Call Subclass_AddMsg(lngWindow, WM_TIMER)
            Call Subclass_AddMsg(lngWindow, WM_PAINT)
            Call Subclass_Initialize(GetParent(lngWindow))
            Call Subclass_AddMsg(GetParent(lngWindow), WM_COMMAND)
            
            ' ComboBox Style is: 1 - Simple Combo (there is no button)
            If Not GetComboBoxButton(lngWindow, lngWindow) Then
               Call Subclass_Initialize(lngWindow)
               Call Subclass_AddMsg(lngWindow, WM_PAINT)
            End If
         End If
      Next 'ctlControl
      
      On Local Error GoTo 0
      Set ctlControl = Nothing
   End If

    '<EhFooter>
    Exit Sub

Initialize_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.Initialize", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub

Private Sub UserControl_Initialize()
    '<EhHeader>
    On Error GoTo UserControl_Initialize_Err
    '</EhHeader>

   IsThemed = CheckIsThemed

    '<EhFooter>
    Exit Sub

UserControl_Initialize_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.UserControl_Initialize", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub

Private Sub UserControl_InitProperties()
    '<EhHeader>
    On Error GoTo UserControl_InitProperties_Err
    '</EhHeader>

   DefaultBorderColor = GetDefaultBorderColor
   m_ComboBoxBorderColor = DefaultBorderColor
   m_DriveListBoxBorderColor = DefaultBorderColor
   m_ImageComboBorderColor = DefaultBorderColor

    '<EhFooter>
    Exit Sub

UserControl_InitProperties_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.UserControl_InitProperties", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error GoTo UserControl_ReadProperties_Err
    '</EhHeader>

   With PropBag
      DefaultBorderColor = GetDefaultBorderColor
      m_BorderColorStyle = .ReadProperty("BorderColorStyle", ThemeColors)
      picButtonState.item(Disabled).Picture = .ReadProperty("ButtonDisabled", Nothing)
      picButtonState.item(Normal).Picture = .ReadProperty("ButtonNormal", Nothing)
      picButtonState.item(Over).Picture = .ReadProperty("ButtonOver", Nothing)
      picButtonState.item(Pressed).Picture = .ReadProperty("ButtonPressed", Nothing)
      m_ButtonThemeType = .ReadProperty("ButtonThemeType", Windows)
      m_ComboBoxBorderColor = .ReadProperty("ComboBoxBorderColor", DefaultBorderColor)
      m_DriveListBoxBorderColor = .ReadProperty("DriveListBoxBorderColor", DefaultBorderColor)
      m_ImageComboBorderColor = .ReadProperty("ImageComboBorderColor", DefaultBorderColor)
   End With
   
   Call CheckButtonThemeType
   
   If IsThemedWindows Then
      ' First subclass the Parent of the UserControl
      ' So we can catch the controls when the Parent activate
      Call Subclass_Initialize(Parent.hWnd)
      Call Subclass_AddMsg(Parent.hWnd, WM_ACTIVATE)
      Call Subclass_AddMsg(Parent.hWnd, WM_THEMECHANGED)
   End If

    '<EhFooter>
    Exit Sub

UserControl_ReadProperties_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.UserControl_ReadProperties", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub

Private Sub UserControl_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

Static blnBusy As Boolean

   If blnBusy Then Exit Sub
   
   blnBusy = True
   Width = picImage.Width
   Height = picImage.Height
   blnBusy = False

End Sub

Private Sub UserControl_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

   On Local Error GoTo ExitSub
   
   Call Subclass_Terminate
   
ExitSub:
   On Local Error GoTo 0
   Erase SubclassData

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error GoTo UserControl_WriteProperties_Err
    '</EhHeader>

   With PropBag
      .WriteProperty "BorderColorStyle", m_BorderColorStyle, ThemeColors
      .WriteProperty "ButtonDisabled", picButtonState.item(Disabled).Picture, Nothing
      .WriteProperty "ButtonNormal", picButtonState.item(Normal).Picture, Nothing
      .WriteProperty "ButtonOver", picButtonState.item(Over).Picture, Nothing
      .WriteProperty "ButtonPressed", picButtonState.item(Pressed).Picture, Nothing
      .WriteProperty "ButtonThemeType", m_ButtonThemeType, Windows
      .WriteProperty "ComboBoxBorderColor", m_ComboBoxBorderColor, GetDefaultBorderColor
      .WriteProperty "DriveListBoxBorderColor", m_DriveListBoxBorderColor, GetDefaultBorderColor
      .WriteProperty "ImageComboBorderColor", m_ImageComboBorderColor, GetDefaultBorderColor
   End With

    '<EhFooter>
    Exit Sub

UserControl_WriteProperties_Err:
    Err.Raise vbObjectError + 100, _
              "prjFrecRodamientos.ThemedComboBox.UserControl_WriteProperties", _
              "ThemedComboBox falla de componente"
    '</EhFooter>
End Sub
