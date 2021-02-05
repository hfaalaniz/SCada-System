Attribute VB_Name = "mdlInsertarVarios"
Option Explicit
'-------------------------------------------------------------
    ' Declaraciones Api
'-------------------------------------------------------------
'Recupera el Hwnd de un menú
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'Elimina el menú de una aplicación
Private Declare Function DeleteMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Recupera la cantidad de Item de menúes para saber cuantos hay que eliminar
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
'Redibuja - repinta la barra de menú luego de eliminarlo
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
'Para cerrar-finalizar una apicación abierta por medio de su HWND
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Api: busca el Handle del programa
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
' función Api SetParent
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndParent As Long) As Long
' Declaración de la función Api ShowWindow
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'Esta función recupera el ancho y alto del área  cliente de la ventana en pixeles
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
' Estas tres funciones es para eliminar la barra de título   del programa que se va a incrustar
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'-------------------------------------------------------------
    ' Constantes
'-------------------------------------------------------------
'Constante para ShowWindow - para maximizar la ventana
Const SHOWMAXIMIZED_eSW = 3&
'Constante para usar con el Api DeleteMenu
Const MF_BYPOSITION = &H400&
Const MF_REMOVE = &H1000&
'Constante para usar con el Api SendMessage para cerrar  la aplicación ( en este caso La calculadora )
Const SC_CLOSE = &HF060&
Const WM_SYSCOMMAND = &H112
'Constante para usar con GetWindowLong y SetWindowLong
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
'Constantes para SetWindowPos
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
'Para usar con el Api GetClientRect
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'-------------------------------------------------------------
    ' Variables
'-------------------------------------------------------------
Dim APP_Rect As RECT
' Mantiene el Handle del programa
Public El_Hwnd_Programa As Long
'-------------------------------------------------------------
    ' Pocedimientos y funciones
'-------------------------------------------------------------
'Elimina y reestablece la barra de título de una ventana
'El primer parámetro es el Hwnd de la misma
 
Sub Quitar_Barra_Titulo(ByVal El_Hwnd_Programa As Long, ByVal Quitar As Boolean)
Dim El_Estilo As Long
    'Almacena en la variable el estilo actual
    El_Estilo = GetWindowLong(El_Hwnd_Programa, GWL_STYLE)
    If Not Quitar Then
        El_Estilo = El_Estilo Or WS_CAPTION
    Else
        El_Estilo = El_Estilo And Not WS_CAPTION
    End If
    'Aplica el nuevo estilo
    SetWindowLong El_Hwnd_Programa, GWL_STYLE, El_Estilo
    
    SetWindowPos El_Hwnd_Programa, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER
End Sub

'Elimina el menú de una ventana específica
 Sub Eliminar_Menu(El_Hwnd_Programa As Long)
    Dim hwnd_Menu As Long
    Dim n_Menu As Long
    Dim i As Integer
    ' Recuper el hwnd del menu del programa
    hwnd_Menu = GetMenu(El_Hwnd_Programa)
    If hwnd_Menu Then
        'cantidad de menúes
        n_Menu = GetMenuItemCount(hwnd_Menu)
        If n_Menu Then
        'Recorre todos los menú y los elimina
        For i = 1 To n_Menu
           Call DeleteMenu(hwnd_Menu, 0, MF_BYPOSITION Or MF_REMOVE)
        Next
        'Repinta la barra de menú
        Call DrawMenuBar(El_Hwnd_Programa)
        End If
    End If
End Sub

'Cierra
Sub Cerrar_Programa(El_Hwnd_Programa As Long)
'Cierra el programa abierto, en este caso la calculadora
Call SendMessage(El_Hwnd_Programa, WM_SYSCOMMAND, SC_CLOSE, ByVal 0&)
End Sub

'mete la ventana en el contenedor
Sub Incrustar_calculadora(Path_programa As String, Contenedor As Object, Titulo_Ventana As String, El_Form As Form)
Dim OLD_Scale As Integer
'almacena la escala para reestablecerla luego
OLD_Scale = El_Form.ScaleMode
El_Form.ScaleMode = vbPixels
If El_Hwnd_Programa = 0 Then
    'Abre el programa
    Shell Path_programa, vbMinimizedNoFocus
    DoEvents
    'Handle de la aplicación
    El_Hwnd_Programa = FindWindow(vbNullString, Titulo_Ventana)
    'REcupera el ancho y alto del área cliente
    Call GetClientRect(El_Hwnd_Programa, APP_Rect)
    'Redimensiona el picturebox al ancho y alto del programa
    Contenedor.Width = (APP_Rect.Right - APP_Rect.Left)
    Contenedor.Height = (APP_Rect.Bottom - APP_Rect.Top)
    Call ShowWindow(El_Hwnd_Programa, vbHide)
    'Elimina la barra de título, los menúes y lo incrusta
    Call Quitar_Barra_Titulo(El_Hwnd_Programa, True)
    Call Eliminar_Menu(El_Hwnd_Programa)
    Call Incrustar(El_Hwnd_Programa, Contenedor)
End If
El_Form.ScaleMode = OLD_Scale
End Sub

Private Sub Incrustar(h_Programa As Long, el_Contenedor As Object)
Dim Ret As Long
'Lo metemos dentro del Picture1
Call SetParent(h_Programa, el_Contenedor.hwnd)
'Maximizamos la ventana incrustada dentro del contenedor, mediante el Api showWindow, pasándole la constante SHOWMAXIMIZED_eSW
Ret = ShowWindow(h_Programa, SHOWMAXIMIZED_eSW)
End Sub

' Libera la ventana pasándole en el segundo parámetro el valor 0 y la cierra
Sub Liberar_Programa(el_Hwnd As Long)
If el_Hwnd <> 0 Then
    ' Libera el programa
    Call SetParent(el_Hwnd, 0)
    'Lo cierra
    Call Cerrar_Programa(El_Hwnd_Programa)
    El_Hwnd_Programa = 0
End If
End Sub

'Codigo para posicionar un progressbar en un statusbar sin setparent
Sub Posicionar(ProgressBar As ProgressBar, StausBar As StatusBar, IndicePanel As Integer)
  Dim Marg As Single
  Dim Panel As Panel
  Dim PosY As Single
  Marg = ProgressBar.Parent.ScaleY(2, vbPixels, vbTwips)
  ' referencia al panel indicado en el parámetro
  Set Panel = StausBar.Panels(IndicePanel)
  ' estilo
  Panel.Bevel = sbrNoBevel
  ' posición top
  PosY = ProgressBar.Parent.ScaleHeight - StausBar.Height
  ' posiciona la barra
  ProgressBar.Move Panel.Left, (PosY + Marg), Panel.Width, (StausBar.Height - Marg)
 ' trae la barra al frente
  ProgressBar.ZOrder
End Sub
'------------colocar en el evento resize del form---------------
    ' pone el progress en el panel número 1
    'Call Posicionar(ProgressBar1, StatusBar1, 1)
'------colocar en el evento form_load del formulario------------
'    With ProgressBar1
'        .value = 50
'        .Max = 100
'        .Scrolling = ccScrollingSmooth
'    End With
'    With StatusBar1
'        ' agrega tres paneles
'        .Panels.Clear
'        .Panels.Add 1
'        .Panels.Add 2, , " Panel 2"
'        .Panels.Add 3, , " Panel 3"
'        .Panels(3).AutoSize = sbrSpring
'    End With

'Posicionar progressbar en statusbar con setparent
Sub PosicionarStBar()
        'Le pasamos a SetParent el HWND de la barra de progreso  y el Statusbar
        ''SetParent ProgressBar1.hwnd, StatusBar1.hwnd
        ''ProgressBar1.Top = 55
        ' Posición izquierda del Progressbar
        ''ProgressBar1.Left = StatusBar1.Panels(5).Left
        'Ancho y alto del ProgressBar igual al panel del Status B
        ''ProgressBar1.Width = StatusBar1.Panels(5).Width
        ''ProgressBar1.Height = StatusBar1.Height - 90
        'Esta Línea es para que se vea llena la barra ( Opcional )
        ''ProgressBar1.value = ProgressBar1.Max
End Sub
'Esto se coloca en el formaload
'    'Agregamos algunos paneles al control Statusbar con un texto
'    StatusBar1.Panels.Add 1, , " Panel 1 "
'    StatusBar1.Panels.Add 2, , " Panel 2 "
'    StatusBar1.Panels.Add 3, , " Panel 3 "
'    StatusBar1.Panels.Add 4, , " Panel 4 "
'
'    Command1.Caption = " Incrustar Progressbar en la barra "


