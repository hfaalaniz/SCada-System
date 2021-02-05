Attribute VB_Name = "AlwaysOnTop"
Option Explicit
'============================================================
' Funcion que se usa de Always on Top con Siempre_Encima
'============================================================
Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const flags = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1 'Constante que uso para activar la Funcion
                        'Siempre_Encima
Const HWND_NOTOPMOST = 1  'Constante que uso para desactivar la
                          'Funcion siempre_Encima

Function Siempre_Encima(frm As Form, Encima As Integer)
   'On Error GoTo Err_Proc

    If Encima Then
        Encima = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    Else
        Encima = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
    End If

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "AlwaysOnTop", "Siempre_Encima"
   Err.Clear
   Resume Exit_Proc

End Function
