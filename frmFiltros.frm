VERSION 5.00
Begin VB.Form frmFiltros 
   Caption         =   "Aplicacion de filtro Chevishev"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3555
      Left            =   90
      ScaleHeight     =   3495
      ScaleWidth      =   11085
      TabIndex        =   9
      Top             =   3630
      Width           =   11145
   End
   Begin VB.CommandButton cmdCalFiltro 
      Caption         =   "Calcular Filtro"
      Height          =   315
      Left            =   6030
      TabIndex        =   8
      Top             =   2910
      Width           =   1335
   End
   Begin VB.TextBox txtNP 
      Height          =   345
      Left            =   6000
      TabIndex        =   3
      Top             =   2130
      Width           =   1275
   End
   Begin VB.TextBox txtPR 
      Height          =   345
      Left            =   6000
      TabIndex        =   2
      Top             =   1740
      Width           =   1275
   End
   Begin VB.TextBox txtLH 
      Height          =   345
      Left            =   6000
      TabIndex        =   1
      Top             =   1350
      Width           =   1275
   End
   Begin VB.TextBox txtFC 
      Height          =   345
      Left            =   6000
      TabIndex        =   0
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label lblB1 
      Caption         =   "B1"
      Height          =   285
      Left            =   8790
      TabIndex        =   11
      Top             =   2730
      Width           =   705
   End
   Begin VB.Label lblA1 
      Caption         =   "A1"
      Height          =   285
      Left            =   8790
      TabIndex        =   10
      Top             =   2370
      Width           =   705
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Enter number of poles (2,4,...20):"
      Height          =   195
      Left            =   3675
      TabIndex        =   7
      Top             =   2160
      Width           =   2310
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Enter percent ripple    (0 to 29):"
      Height          =   195
      Left            =   3795
      TabIndex        =   6
      Top             =   1770
      Width           =   2190
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Enter  0 for LP,  1 for HP filter:"
      Height          =   195
      Left            =   3870
      TabIndex        =   5
      Top             =   1380
      Width           =   2115
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Enter cutoff frequency  (0 to .5):"
      Height          =   195
      Left            =   3735
      TabIndex        =   4
      Top             =   990
      Width           =   2250
   End
End
Attribute VB_Name = "frmFiltros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i, p As Single
Dim A0, A1, A2, B1, B2, SA, SB As Long
Dim GAIN As Long
'Private Const PI = 3.14159265
'CHEBYSHEV FILTER-  COEFFICIENT CALCULATION
                         'INITIALIZE VARIABLES
 Dim a(22) As Single       'holds the "a" coefficients
 Dim b(22) As Single       'holds the "b" coefficients
 Dim TA(22) As Single      'internal use for combining stages
 Dim TB(22) As Single      'internal use for combining stages
 Sub Z_Transfor()
 'THIS SUBROUTINE IS CALLED FROM FIG. 33-7,  LINE 340
 'Variables entering subroutine: PI, txtFC, txtLH, txtPR, HP, P
 'Variables exiting subroutine:   A0, A1, A2, B1, B2
 'Variables used internally:
 Dim RP, IP, ES, VX, KX, T, W, m, D, k, X0, X1, X2, Y1, Y2 As Long
 
 ' 'Calculate pole location on unit circle
 RP = -Cos(PI / (txtNP * 2) + (p - 1) * PI / txtNP)
 IP = Sin(PI / (txtNP * 2) + (p - 1) * PI / txtNP)
 ' 'Urdimbre de un círculo a una elipse
1210:
If txtPR = 0 Then GoTo 1210
    ES = Sqr((100 / (100 - txtPR)) ^ 2 - 1)
    VX = (1 / txtNP) * Log((1 / ES) + Sqr((1 / ES ^ 2) + 1))
    KX = (1 / txtNP) * Log((1 / ES) + Sqr((1 / ES ^ 2) - 1))
    KX = (Exp(KX) + Exp(-KX)) / 2
    RP = RP * ((Exp(VX) - Exp(-VX)) / 2) / KX
    IP = IP * ((Exp(VX) + Exp(-VX)) / 2) / KX
    '             's-domain to z-domain conversion
    T = 2 * Tan(1 / 2)
    W = 2 * PI * txtFC
    m = RP ^ 2 + IP ^ 2
    D = 4 - 4 * RP * T + m * T ^ 2
    X0 = T ^ 2 / D
    X1 = 2 * T ^ 2 / D
    X2 = T ^ 2 / D
    Y1 = (8 - 2 * m * T ^ 2) / D
    Y2 = (-4 - 4 * RP * T - m * T ^ 2) / D
    ' 'LP TO LP, or LP TO HP
    If txtLH = 1 Then k = -Cos(W / 2 + 1 / 2) / Cos(W / 2 - 1 / 2)
    If txtLH = 0 Then k = Sin(1 / 2 - W / 2) / Sin(1 / 2 + W / 2)
    D = 1 + Y1 * k - Y2 * k ^ 2
    A0 = (X0 - X1 * k + X2 * k ^ 2) / D
    A1 = (-2 * X0 * k + X1 + X1 * k ^ 2 - 2 * X2 * k) / D
    A2 = (X0 * k ^ 2 - X1 * k + X2) / D
    B1 = (2 * k + Y1 + Y1 * k ^ 2 - 2 * Y2 * k) / D
    B2 = (-k ^ 2 - Y1 * k + Y2) / D
    If txtLH = 1 Then A1 = -A1
    If txtLH = 1 Then B1 = -B1
    
End Sub

Sub FiltroCHEBYSHEV()
 For i = 0 To 22
   a(i) = 0
   b(i) = 0
 Next i
 a(2) = 1
 b(2) = 1
 'PI = 3.14159265
                           'ENTER THE FILTER PARAMETERS
 'INPUT "Enter cutoff frequency  (0 to .5): ", txtFC
 'INPUT "Enter  0 for LP,  1 for HP filter: ", txtLH
 'INPUT "Enter percent ripple    (0 to 29):  ", txtPR
 'INPUT "Enter number of poles (2,4,...20): ", txtNP
 For p = 1 To txtNP / 2  'LOOP FOR EACH POLE-ZERO PAIR
   Call Z_Transfor   'GoSub 1000     'The subroutine in Fig. 33-8
   For i = 0 To 22    'Add coefficients to the cascade
     TA(i) = a(i)
     TB(i) = b(i)
   Next i
   For i = 2 To 22
     a(i) = A0 * TA(i) + A1 * TA(i - 1) + A2 * TA(i - 2)
     b(i) = TB(i) - B1 * TB(i - 1) - B2 * TB(i - 2)
   Next i
 Next p
 b(2) = 0            'Finish combining coefficients
 For i = 0 To 20
   a(i) = a(i + 2)
   b(i) = -b(i + 2)
 Next i
 '
 SA = 0               'NORMALIZE THE GAIN
 SB = 0
 For i = 0 To 20
   If txtLH = 0 Then SA = SA + a(i)
   If txtLH = 0 Then SB = SB + b(i)
   If txtLH = 1 Then SA = SA + a(i) * (-1) ^ i
   If txtLH = 1 Then SB = SB + b(i) * (-1) ^ i
 Next i
 '
 GAIN = SA / (1 - SB)
 lblA1.Caption = Val(GAIN)
 For i = 0 To 20
    a(i) = a(i) / GAIN
 Next i
 ' 'The final recursion coefficients are
' End 'in A( ) and B( )
End Sub

Private Sub cmdCalFiltro_Click()
    Call FiltroCHEBYSHEV
End Sub
