VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   276
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   498
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4125
      Left            =   0
      ScaleHeight     =   275
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   499
      TabIndex        =   0
      Top             =   0
      Width           =   7485
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3585
         Left            =   1215
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   3585
         ScaleWidth      =   5205
         TabIndex        =   3
         Top             =   270
         Width           =   5205
      End
      Begin VB.PictureBox pM 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3795
         Left            =   150
         ScaleHeight     =   253
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   479
         TabIndex        =   1
         Top             =   180
         Width           =   7185
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   4320
            Top             =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Versón DEMO"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   30
            TabIndex        =   2
            Top             =   30
            Visible         =   0   'False
            Width           =   1020
         End
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variable tipo Flag que indica cuando se cumplíó el tiempo para descargar la pantalla de presentación
Public Listo As Boolean
'****************************************************
' Valor en segundos para la duración del Splash
'****************************************************
Private Const DURACION As Long = 2

Private Const SRCCOPY = &HCC0020
Private Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "User32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function LockWindowUpdate Lib "User32" (ByVal hwndLock As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Private Declare Function SetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Type rgbPixel
    r As Byte
    g As Byte
    b As Byte
End Type

Private Type Shadow
    Draw As Boolean
    Intensity As Single
    Length As Long
    Shift As PointAPI
End Type

Private Type Rounded_Rectangle
    Start As PointAPI
    Diamension As PointAPI
    Round_Radius As Long
    DropShadow As Shadow
    BackColor As rgbPixel
End Type

Public Enum Form_style_t
    About = 0
    Splash = 1
End Enum

Private Form_style_r As Form_style_t

Dim TopWindow As Long
Dim Shadow_radius As Long

Private Sub GetRGB(ByVal Col As Long, ByRef Red, ByRef Green, ByRef Blue)
  Red = Col Mod 256
  Green = ((Col And &HFF00&) \ 256&) Mod 256&
  Blue = (Col And &HFF0000) \ 65536
End Sub

Private Sub RoundedRectTrans(p As PictureBox, r As Rounded_Rectangle)
    'This sub draws a rounded rectangle
    Dim X As Long, Y As Long, I As Long, Rradius As Long
    Dim Sx As Long, Ex As Long
    Dim c As Single, m1c As Single, cr As Long, cg As Long, cb As Long, cr2 As Long, cg2 As Long, cb2 As Long
    Dim tV1 As Single, tV2 As Single, tV As Single
    Dim X1 As Long, X2 As Long, X3 As Long, X4 As Long
    Dim Y1 As Long, Y2 As Long, Y3 As Long, Y4 As Long
    Dim BkR As Long, BkB As Long, BkG As Long
    If r.DropShadow.Draw Then
        Rradius = r.Round_Radius + r.DropShadow.Length
        
        For Y = r.Start.Y + Rradius - r.DropShadow.Length + 1 To r.Start.Y + r.Diamension.Y - Rradius + r.DropShadow.Length - 1
            'Left
            tV = (r.Start.X - r.DropShadow.Length)
            tV2 = r.Start.X - r.DropShadow.Length
            For X = tV2 To r.Start.X
                c = (X - tV) / r.DropShadow.Length
                m1c = 1 - c
                GetRGB GetPixel(p.hDC, X + r.DropShadow.Shift.X, Y + r.DropShadow.Shift.Y), BkR, BkG, BkB
                
                cr = BkR * m1c + r.BackColor.r * c
                cg = BkG * m1c + r.BackColor.g * c
                cb = BkB * m1c + r.BackColor.b * c
                SetPixel p.hDC, X + r.DropShadow.Shift.X, Y + r.DropShadow.Shift.Y, RGB(cr, cg, cb)
            Next X
            'Right
            tV1 = (r.Start.X + r.Diamension.X)
            tV2 = r.Start.X + r.Diamension.X + r.DropShadow.Length
            For X = r.Start.X + r.Diamension.X To tV2
                m1c = (X - tV1) / r.DropShadow.Length
                c = 1 - m1c
                GetRGB GetPixel(p.hDC, X + r.DropShadow.Shift.X, Y + r.DropShadow.Shift.Y), BkR, BkG, BkB
                
                cr = BkR * m1c + r.BackColor.r * c
                cg = BkG * m1c + r.BackColor.g * c
                cb = BkB * m1c + r.BackColor.b * c
                SetPixel p.hDC, X + r.DropShadow.Shift.X, Y + r.DropShadow.Shift.Y, RGB(cr, cg, cb)

            Next X
        Next Y

        Rradius = r.Round_Radius + r.DropShadow.Length
        X2 = r.DropShadow.Shift.X
        X1 = r.Start.X + r.DropShadow.Shift.X + r.Diamension.X + r.Start.X
        For Y = 0 To Rradius
            X = Sqr(Rradius * Rradius - Y * Y) + r.DropShadow.Length
            Ex = IIf(r.Round_Radius >= Y, r.Start.X + r.Round_Radius - Sqr(Abs(r.Round_Radius * r.Round_Radius - Y * Y)), r.Start.X + r.Round_Radius)
            tV1 = -r.Start.X - Rradius + r.DropShadow.Length
            tV2 = (r.Start.X + Rradius - X)
            Y1 = r.Start.Y - Y + Rradius + r.DropShadow.Shift.Y - r.DropShadow.Length
            Y2 = r.Start.Y + Y - Rradius + r.Diamension.Y + r.DropShadow.Shift.Y + r.DropShadow.Length


           For I = tV2 To Ex
                m1c = (Sqr((I + tV1) * (I + tV1) + Y * Y) - r.Round_Radius) / r.DropShadow.Length
                c = 1 - m1c

                GetRGB GetPixel(p.hDC, I + X2, Y1), BkR, BkG, BkB
                
                cr = BkR * m1c + r.BackColor.r * c
                cg = BkG * m1c + r.BackColor.g * c
                cb = BkB * m1c + r.BackColor.b * c
                
                cr = Abs(cr)
                cg = Abs(cg)
                cb = Abs(cb)
                
                'Top left
                SetPixel p.hDC, I + X2, Y1, RGB(cr, cg, cb)
                'Bottom left
                GetRGB GetPixel(p.hDC, I + X2, Y2), BkR, BkG, BkB
                
                cr = BkR * m1c + r.BackColor.r * c
                cg = BkG * m1c + r.BackColor.g * c
                cb = BkB * m1c + r.BackColor.b * c
                
                cr = Abs(cr)
                cg = Abs(cg)
                cb = Abs(cb)
                
                
                SetPixel p.hDC, I + X2, Y2, RGB(cr, cg, cb)
                
                
                GetRGB GetPixel(p.hDC, -I + X1, Y1), BkR, BkG, BkB
                
                cr = BkR * m1c + r.BackColor.r * c
                cg = BkG * m1c + r.BackColor.g * c
                cb = BkB * m1c + r.BackColor.b * c
                
                cr = Abs(cr)
                cg = Abs(cg)
                cb = Abs(cb)
                
                'Top right
                SetPixel p.hDC, -I + X1, Y1, RGB(cr, cg, cb)
                
                GetRGB GetPixel(p.hDC, -I + X1, Y2), BkR, BkG, BkB
                
                cr = BkR * m1c + r.BackColor.r * c
                cg = BkG * m1c + r.BackColor.g * c
                cb = BkB * m1c + r.BackColor.b * c
                
                cr = Abs(cr)
                cg = Abs(cg)
                cb = Abs(cb)
                SetPixel p.hDC, -I + X1, Y2, RGB(cr, cg, cb)
                
            Next I
        Next Y
        
        X1 = Rradius + r.Start.X - r.DropShadow.Length + r.DropShadow.Shift.X
       
        For Y = 0 To r.DropShadow.Length
            c = Y / r.DropShadow.Length
            m1c = 1 - c
            
            tV1 = r.Diamension.X - (r.Round_Radius * 2)
            Y1 = Y + r.Start.Y - r.DropShadow.Length + r.DropShadow.Shift.Y
            Y2 = -Y + r.Start.Y + r.Diamension.Y + r.DropShadow.Length + r.DropShadow.Shift.Y

            For X = 1 To tV1 - 1
                GetRGB GetPixel(p.hDC, X + X1, Y1), BkR, BkG, BkB
                cr = BkR * m1c + r.BackColor.r * c
                cg = BkG * m1c + r.BackColor.g * c
                cb = BkB * m1c + r.BackColor.b * c
                
                GetRGB GetPixel(p.hDC, X + X1, Y2), BkR, BkG, BkB
                cr2 = BkR * m1c + r.BackColor.r * c
                cg2 = BkG * m1c + r.BackColor.g * c
                cb2 = BkB * m1c + r.BackColor.b * c
            
                'top
                SetPixel p.hDC, X + X1, Y1, RGB(cr, cg, cb)
                'bottom
                SetPixel p.hDC, X + X1, Y2, RGB(cr2, cg2, cb2)
            Next X
        Next Y
    End If
End Sub

Private Sub MakeShadow(Optional Shadow_size = 15, Optional Shadow_color As Long = 0, Optional Shadow_shiftX As Long = 0, Optional Shadow_shiftY As Long = 0)
    Dim xSrc As Long, ySrc As Long, hSrcDC As Long
    Dim r As Rounded_Rectangle
    
    'Get screenshot
    xSrc = (Screen.Width / 2) / Screen.TwipsPerPixelX - Me.ScaleWidth / 2
    ySrc = (Screen.Height / 2) / Screen.TwipsPerPixelY - Me.ScaleHeight / 2
    hSrcDC = GetDC(0)
    BitBlt p.hDC, 0, 0, p.ScaleWidth, p.ScaleHeight, hSrcDC, xSrc, ySrc, vbSrcCopy
    ReleaseDC 0, GetDC(0)
    'Set Shadow properties
    With r
        .DropShadow.Shift.X = Shadow_shiftX
        .DropShadow.Shift.Y = Shadow_shiftY
        .Diamension.X = p.ScaleWidth - Shadow_size * 2 - 1 - .DropShadow.Shift.X
        .Diamension.Y = p.ScaleHeight - Shadow_size * 2 - 1 - .DropShadow.Shift.Y
        .DropShadow.Length = Shadow_size
        .Start.X = Shadow_size
        .Start.Y = Shadow_size
        .DropShadow.Draw = True
        GetRGB Shadow_color, .BackColor.r, .BackColor.g, .BackColor.b
    End With
    'Draw shadow
    RoundedRectTrans p, r
    p.Refresh
End Sub
Private Sub UpdateShadow()
    
    Dim Old_top As Long, Old_left As Long
    
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    Me.Show
    Old_top = Me.Top: Old_left = Me.Left
    'Hide the form from the screen setting me.visible to false
    Me.Top = Screen.Height + 100: Me.Left = Screen.Width + 100
    'Refresh screen
    DoEvents
    'Draw the shadow
    MakeShadow (p.ScaleWidth - pM.ScaleWidth) / 2, RGB(50, 50, 40), 8, 10
    'Reset form position
    Me.Left = Old_left: Me.Top = Old_top
    'Start timer
    TopWindow = Me.hWnd
End Sub

Private Sub Form_Load()
  Dim El_Tiempo As Long
      Shadow_radius = 16
    p.Width = pM.Width + Shadow_radius * 2
    p.Height = pM.Height + Shadow_radius * 2
    pM.Move p.ScaleWidth / 2 - pM.ScaleWidth / 2, p.ScaleHeight / 2 - pM.ScaleHeight / 2
    
    Me.Width = p.ScaleWidth * Screen.TwipsPerPixelX
    Me.Height = p.ScaleHeight * Screen.TwipsPerPixelY
    
    UpdateShadow
  ' Muestra la ventana y le establece la animación
   'On Error GoTo Err_Proc
  Call Animar(Me, 500, AW_CENTER Or AW_ACTIVATE)
  'lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
  'lblProductName.Caption = App.Title
  'lblCompanyProduct.Caption = App.CompanyName
  'Almacena el punto de partida para hacer el retardo en segundos
  El_Tiempo = Timer + DURACION
  ' ... Espera
  Do While Timer < El_Tiempo
     DoEvents
  Loop
  ' Se cierra desde SubMain
  Listo = True
End Sub

Private Sub Timer1_Timer()
    Dim CurTopWindow As Long
    CurTopWindow = GetForegroundWindow
    
    'Check to see if a new window has blocked this form
    If CurTopWindow <> TopWindow Then
        If Form_style_r = Splash Then
            'If splash, then must redraw the window after the user has released the mouse button
            If GetAsyncKeyState(1) = 0 Then
                UpdateShadow
            End If
        Else
            'if about box then unload
            Unload Me
        End If
    End If
End Sub

Public Property Get Form_style() As Form_style_t
    Form_style = Form_style_r
End Property

Public Property Let Form_style(ByVal vNewValue As Form_style_t)
    Form_style_r = vNewValue
End Property

