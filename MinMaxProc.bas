Attribute VB_Name = "MinMaxProc"
'------------------------------------------------------------------------
'Esta es una ventana procedimiento de restringir el tamaño de un Ventana.
'------------------------------------------------------------------------
Option Explicit

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Proc As Long
Public Const GWL_WNDPROC = (-4)

Private Const WM_GETMINMAXINFO = &H24

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal ByteLength As Long)

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type


Function WindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Static MinMax As MINMAXINFO
    
   'On Error GoTo Err_Proc

    Select Case Msg
        Case WM_GETMINMAXINFO
            Call MoveMemory(MinMax, ByVal lParam, Len(MinMax))
            MinMax.ptMinTrackSize.x = MinWidth    'Base.MinWidth
            MinMax.ptMinTrackSize.y = MinHeight   'Base.MinHeight
            Call MoveMemory(ByVal lParam, MinMax, Len(MinMax))
        Case Else
            WindowProc = CallWindowProc(Proc, hwnd, Msg, wParam, lParam)
    End Select

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "MinMaxProc", "WindowProc"
   Err.Clear
   Resume Exit_Proc

End Function


