Attribute VB_Name = "mdlMain"
Option Explicit
Public Tooltips As New Collection
Public Tooltip   As cToolTip
'Variable tipo Flag que indica cuando se cumplíó el tiempo para descargar la pantalla de presentación
Public Listo As Boolean
Public AutorizadoA As String

Const GWL_WNDPROC = (-4)
' Declaraciones del Api
''''''''''''''''''''''''''''''''''''''''''
Private Declare Function SendMessageByString& Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)
' función que deshabilita el repintado de una ventana en windows
Private Declare Function LockWindowUpdate& Lib "User32" (ByVal hwndLock As Long)
' variables y constantes
Private Const CB_ADDSTRING& = &H143
Private Const LB_ADDSTRING As Long = &H180
'*************************************************************************
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const PROP_PREVPROC = "PrevProc"
Const PROP_FORM = "FormObject"
Private Declare Function SetProp Lib "User32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "User32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "User32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal DestL As Long)
Public Declare Function SendMessageLong Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
'Flag para la tecla BackSpace
Public KeyRetroceso As Boolean
Global PI As Double

Const WM_PRINTCLIENT = &H318

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function apiOleTranslateColor Lib "oleaut32" Alias "OleTranslateColor" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Enum AnimateWindowFlags
    AW_HOR_POSITIVE = &H1
    AW_HOR_NEGATIVE = &H2
    AW_VER_POSITIVE = &H4
    AW_VER_NEGATIVE = &H8
    AW_CENTER = &H10
    AW_HIDE = &H10000
    AW_ACTIVATE = &H20000
    AW_SLIDE = &H40000
    AW_BLEND = &H80000
End Enum
Private Declare Function AnimateWindow Lib "User32" (ByVal hWnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal Mul As Long, ByVal Nom As Long, ByVal Den As Long) As Long
Private Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "User32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function FillRect Lib "User32" (ByVal hDC As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' ( Sub Main ) Procedimiento de Inicio
'******************************************************
Private Sub Main()
PI = 3.14159265
  'Carga y visualiza el formulario Splash
   'On Error GoTo Err_Proc
  If App.PrevInstance = True Then MsgBox App.EXEName & " ya esta cargado en su sistema. Saliendo...", vbInformation, "Atención": End
  Load frmSplash
  ' Carga en memoria el formulario principal pero no lo muestra
  'Load frmPrincipal
  ' ..Hasta que no se cumpla el tiempo se visualiza el Splash
  Do
    DoEvents
  Loop Until frmSplash.Listo
  
  Call Animar(frmSplash, 700, AW_BLEND Or AW_HIDE)
  Unload frmSplash
  Set frmSplash = Nothing
  ' Visualiza el Formulario Principal con el efecto de animación desde el centro
  'Call Animar(frmPrincipal, 200, AW_CENTER Or AW_ACTIVATE)  '
  'descarga el Splash con una animación
  frmPrincipal.Show
  'frmRodamientos.Show

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "mdlMain", "Main"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub Animar(ByVal Form As Form, ByVal dwTime As Long, ByVal dwFlags As AnimateWindowFlags)
    ' Set the properties
   'On Error GoTo Err_Proc
    SetProp Form.hWnd, PROP_PREVPROC, GetWindowLong(Form.hWnd, GWL_WNDPROC)
    SetProp Form.hWnd, PROP_FORM, ObjPtr(Form)
    ' Subclass the window
    SetWindowLong Form.hWnd, GWL_WNDPROC, AddressOf AnimateWinProc
    ' Call AnimateWindow API
    AnimateWindow Form.hWnd, dwTime, dwFlags
    ' Unsubclass the window
    SetWindowLong Form.hWnd, GWL_WNDPROC, GetProp(Form.hWnd, PROP_PREVPROC)
    ' Remove the properties
    RemoveProp Form.hWnd, PROP_FORM
    RemoveProp Form.hWnd, PROP_PREVPROC
    ' Refresh the form
    Form.Refresh
Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "mdlMain", "Animar"
   Err.Clear
   Resume Exit_Proc

End Sub

' AnimateWinProc
'
' Window procedure for AnimateWindow
' ***************************************************************
Private Function AnimateWinProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lPrevProc As Long
    Dim lForm As Long
    Dim oForm As Form
    ' Get the previous WinProc pointer
   'On Error GoTo Err_Proc
    lPrevProc = GetProp(hWnd, PROP_PREVPROC)
    ' Get the form object
    lForm = GetProp(hWnd, PROP_FORM)
    MoveMemory oForm, lForm, 4&

    Select Case Msg
        Case WM_PRINTCLIENT
           Dim tRect As Rect
           Dim hBr As Long
            
            ' Get the window client size
            GetClientRect hWnd, tRect

            ' Create a brush with the
            ' form background color
            hBr = CreateSolidBrush(OleTranslateColor(oForm.BackColor))

            ' Fill the DC with the
            ' background color
            FillRect wParam, tRect, hBr

            ' Delete the brush
            DeleteObject hBr

            If Not oForm.Picture Is Nothing Then
                Dim lScrDC As Long
                Dim lMemDC As Long
                Dim lPrevBMP As Long

                ' Create a compatible DC
                lScrDC = GetDC(0&)
                lMemDC = CreateCompatibleDC(lScrDC)
                ReleaseDC 0, lScrDC

                ' Select the form picture in the DC
                lPrevBMP = SelectObject(lMemDC, oForm.Picture.Handle)

                ' Draw the picture in the DC
                BitBlt wParam, 0, 0, HM2Pix(oForm.Picture.Width), _
                                HM2Pix(oForm.Picture.Height), _
                                lMemDC, 0, 0, vbSrcCopy

                ' Release the picture
                SelectObject lMemDC, lPrevBMP

                ' Delete the DC
                DeleteDC lMemDC

            End If

        End Select

        ' Release the form object
        MoveMemory oForm, 0&, 4&

        ' Call the original window procedure
        AnimateWinProc = CallWindowProc(lPrevProc, hWnd, Msg, wParam, lParam)


Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "mdlMain", "AnimateWinProc"
   Err.Clear
   Resume Exit_Proc

End Function

Private Function HM2Pix(ByVal value As Long) As Long
    HM2Pix = MulDiv(value, 1440, 2540) / Screen.TwipsPerPixelX
End Function

Private Function OleTranslateColor(ByVal Clr As Long) As Long
    apiOleTranslateColor Clr, 0, OleTranslateColor
End Function
 
' Función que carga el campo en el combobox o list
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Cargar(ElControl As Object, rst As ADODB.Recordset, Campo As String) As Boolean
    Dim Ret As Long
    Dim Mensaje_SendMessage As Long
    On Error GoTo error_function:
    ' verifica que el recordset contenga un conjunto de registros
    If rst.BOF And rst.EOF Then
        MsgBox " No hay registros para agregar", vbInformation
        Call LockWindowUpdate(0&)
        ' sale
        Exit Function
    End If
    ElControl.Parent.MousePointer = vbHourglass
    ' Chequea con TypeName el tipo de control enviado como parámetro
    If TypeName(ElControl) = "ComboBox" Then
       Mensaje_SendMessage = CB_ADDSTRING& ' mesanje para SendMessage
    ElseIf TypeName(ElControl) = "ListBox" Then
       Mensaje_SendMessage = LB_ADDSTRING ' mesanje para SendMessage
    End If
    ' deshabilita el repintado del control para que      cargue los datos mas rapidamente
    Call LockWindowUpdate(ElControl.hWnd)
    DoEvents
    ' Posiciona el recordset en el primer registro
    rst.MoveFirst
    ' elimina todo el contenido del combo o listbox( opcional )
    ElControl.Clear
    ' recorre las filas del recordset
    Do Until rst.EOF
        ' chequea que el valor no sea un nulo
        If Not IsNull(rst(Campo).value) Then
            'Agrega el dato en el control con el              mensaje CB_ADDSTRING o LB_ADDSTRING dependiendo del tipo de control
            Ret = SendMessageByString(ElControl.hWnd, Mensaje_SendMessage, 0, rst(Campo).value)
        End If
        ' siguiente registro
        rst.MoveNext
    Loop
    ' selecciona el primer elemento del listado
    If ElControl.ListCount > 0 Then
        ElControl.ListIndex = 0
    End If
    ' vuelve a habilitar el repintado
    Call LockWindowUpdate(0&)
     ' retorno
    Cargar = True
    ' reestablece el puntero del mouse
    ElControl.Parent.MousePointer = vbNormal
    Exit Function
 ' rutina de error
error_function:
    MsgBox Err.Description, vbCritical
    ' En caso de error vuelve a activar el repintado
    Call LockWindowUpdate(0&)
    ElControl.Refresh
    ElControl.Parent.MousePointer = vbNormal
End Function

Function FExsists(strFileName As String) As Boolean ' Does File Already Exsist?
'Dim lpFindFileData As WIN32_FIND_DATA
'Dim hFindFirst As Long
'       hFindFirst = FindFirstFile(strFileName, lpFindFileData)
'              If hFindFirst > 0 Then
'                      FindClose hFindFirst
'                      FExsists = True
'              Else
'                      FExsists = False
'              End If
End Function

Public Sub ShapeCtrl(p As control, Rad As Long, SM As Long)
Dim Reg As Long
' Rad, in this case circular corner radius
' SM 0 for Pixels, 1 for Twips
   If SM = 0 Then
      Reg = CreateRoundRectRgn(0, 0, p.Width, p.Height, Rad, Rad)
   Else
      Reg = CreateRoundRectRgn(1, 1, p.Width \ Screen.TwipsPerPixelX, p.Height \ Screen.TwipsPerPixelY, Rad, Rad)
   End If
   SetWindowRgn p.hWnd, Reg, True
   DeleteObject Reg
End Sub

'***********************************************************************************************
'Rutinas para el manejo de coordenadas, circulos y radio de los graficos polares
'***********************************************************************************************
Private Sub Dib_Rad_Polar(pic As PictureBox, ByVal LargoRadio As Single, ByVal Angulo As Single)
Dim cx, cy, radio, LimiteRadio, CurrentX As Single, CurrentY As Single    ' Declara variable.
Dim xp As Single, yp As Single, rx As Single, ry As Single, rxg As Single, ryg As Single
'PI = 4 * Atn(1)
cx = pic.CurrentX
cy = pic.CurrentY
'El ángulo está en grados
Angulo = Angulo Mod 360
Angulo = Angulo * PI / 180
xp = 0
yp = Abs(LargoRadio)
rx = xp * Cos(Angulo) - yp * Sin(Angulo)
ry = xp * Sin(Angulo) + yp * Cos(Angulo)
rxg = cx + rx
ryg = cy - ry
pic.Line (cx, cy)-(rxg, ryg), &HC0C0C0  'gris oscuro
' si el largo es negativo vuelve a la posición inicial
If LargoRadio < 0 Then
    pic.CurrentX = cx
    pic.CurrentY = cy
End If
End Sub

Private Sub DibujarCirculo(pic As PictureBox)
   Dim cx As Single, cy As Single, radio As Integer, Limite As Integer, ScaleMode      ' Declara variable.
   ScaleMode = 3               ' Establece la escala a píxeles.
   cx = pic.ScaleWidth / 2     ' Establece la posición X.
   cy = pic.ScaleHeight / 2    ' Establece la posición Y.
   If cx > cy Then Limite = cy Else Limite = cx
   
   For radio = 0 To 120 Step 10  ' Establece el radio.
      pic.Circle (cx, cy), radio, &HC0C0C0   'vbBlue
   Next radio
   pic.Circle (cx, cy), 10, vbGreen
End Sub

Private Sub EscalarUnidades(pic As PictureBox)
Dim r As Integer, tangens As Single, cota As Single, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, grad As Single, seno As Single, coseno As Single
r = 8 'radio 8
'escala unidades del círculo polar en el picPIzq----
For grad = 0 To 360
    seno = Sin(grad * PI / 180)
    coseno = Cos(grad * PI / 180)
    X1 = coseno * r
    Y1 = -seno * r
    If grad Mod 5 <> 0 Then
        'grados no son divisibles por 5
        X2 = coseno * (r + 0.2) 'unidades de corta
        Y2 = -seno * (r + 0.2)
        pic.Line (X1, Y1)-(X2, Y2), vbBlue
    Else
        'grados son divisibles por 5
        If grad Mod 10 = 0 Then
            X2 = coseno * (r + 0.6) 'unidades de longitud
            Y2 = -seno * (r + 0.6)
        Else
            X2 = coseno * (r + 0.4) 'unidades de soporte
            Y2 = -seno * (r + 0.4)
        End If
        'unidades de carácter alrededor del círculo
        pic.Line (X1, Y1)-(X2, Y2), vbBlue
    End If
Next grad
End Sub

Private Sub ColocarGrados(pic As PictureBox)
Dim r As Integer, tangens As Single, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, grad As Single, seno As Single, coseno As Single
'círculo gradual ------------
For grad = 0 To 360 Step 10
    r = 9   '8.9    '9   '8.8
    seno = Sin(grad * PI / 180)
    coseno = Cos(grad * PI / 180)
    X2 = coseno * (r + 0.1)  '0.2
    Y2 = -seno * (r + 0.1)   '0.2
    'puntos de partida para los números asignados grado
    Select Case grad
        Case 90
            pic.PSet (X2 - 0.6, Y2 - 0.3), vbBlue   '(X2 - 0.6, Y2 - 0.3), color
        Case 91 To 269
            pic.PSet (X2 - 0.9, Y2 - 0.3), vbBlue     '(X2 - 1, Y2 - 0.3), farve
        Case 270
            pic.PSet (X2 - 0.7, Y2 - 0.3), vbBlue   '(X2 - 0.7, Y2 - 0.3), farve
        Case Else
            pic.PSet (X2 - 0.5, Y2 - 0.3), vbBlue   '(X2 - 0.5, Y2 - 0.3), farve
    End Select
    'coloca los grados alrededor del círculo
    If grad <> 360 Then
        pic.Print grad
    End If
Next grad
End Sub

Sub EscalarPic(pic As PictureBox)
  pic.ScaleMode = 0
  pic.FontSize = 4
  pic.Font = "Segoe UI"
  pic.ScaleTop = -10  ' Set scale for top of grid.
  pic.ScaleLeft = -10 ' Set scale for left of grid.
  pic.ScaleWidth = 20 ' Set scale (-10 to 20).
  pic.ScaleHeight = 20
End Sub

Private Sub EstablecerPlanos(pic As PictureBox)
  pic.ScaleMode = 0
  pic.Height = 5000 '5885 '6885
  pic.Width = 5000  '5885  '6885
  pic.BorderStyle = 1
  pic.BackColor = &HE0E0E0
  pic.FontSize = 4
  pic.Font = "Segoe UI"
  'dimensiones de las coordenadas del sistema
  pic.ScaleTop = -10 '10
  pic.ScaleLeft = -10 '10
  pic.ScaleWidth = 20 '20
  pic.ScaleHeight = 20 '20
End Sub

Private Sub Grilla(pic As PictureBox)
Dim cou, col3, lightgrey, klik, but
'On Error GoTo Err_Proc
pic.Line (-8, 0)-(8, 0), QBColor(8): frmBase.picPIzq.Line (0, -8)-(0, 8), QBColor(8)
pic.Line (-8, 0)-(8, 0), QBColor(8): frmBase.picPDer.Line (0, -8)-(0, 8), QBColor(8)

If lightgrey = 0 Then '<> &HE0E0E0 Then 'dvs. = 0
        lightgrey = QBColor(7)
    End If
For cou = -19 To 19
    pic.Line (-20, cou)-(20, cou), lightgrey 'QBColor(7) '7 er 10
    pic.Line (cou, -20)-(cou, 20), lightgrey 'QBColor(7)
Next cou
lightgrey = QBColor(7)
col3 = 8
pic.Line (-15, -20)-(-15, 40), QBColor(col3) '8 er 10
pic.Line (-10, -20)-(-10, 40), QBColor(col3)
pic.Line (-5, -20)-(-5, 40), QBColor(col3)
pic.Line (5, -20)-(5, 40), QBColor(col3)
pic.Line (10, -20)-(10, 40), QBColor(col3)
pic.Line (15, -20)-(15, 40), QBColor(col3)
pic.Line (-20, -15)-(40, -15), QBColor(col3)
pic.Line (-20, -10)-(40, -10), QBColor(col3)
pic.Line (-20, -5)-(40, -5), QBColor(col3)
pic.Line (-20, 5)-(40, 5), QBColor(col3)
pic.Line (-20, 10)-(40, 10), QBColor(col3)
pic.Line (-20, 15)-(40, 15), QBColor(col3)
pic.Line (-20, 0)-(40, 0), RGB(0, 0, 255) ' Draw horizontal line.
pic.Line (0, -20)-(0, 40), RGB(0, 0, 255) ' Draw vertical line.

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPloteaSeñal", "Grilla"
   Err.Clear
   Resume Exit_Proc
End Sub


