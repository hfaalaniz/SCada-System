VERSION 5.00
Begin VB.Form frmAcercade 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3825
   ClientLeft      =   6975
   ClientTop       =   3435
   ClientWidth     =   9555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versón DEMO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4350
      TabIndex        =   8
      Top             =   90
      Width           =   1020
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Versión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   6150
      TabIndex        =   5
      Top             =   900
      Width           =   885
   End
   Begin VB.Label lblCopyright 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(C) 2012 por Hugo Fabián Alaníz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2130
      TabIndex        =   7
      Top             =   2250
      Width           =   3180
   End
   Begin VB.Label lblCompany 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PLC.Net - hfaalaniz@hotmail.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2190
      TabIndex        =   6
      Top             =   2520
      Width           =   3315
   End
   Begin VB.Label lblPlatform 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Plataforma"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   360
      Left            =   210
      TabIndex        =   4
      Top             =   0
      Width           =   1275
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Autorizado a: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   8370
      TabIndex        =   3
      Top             =   30
      Width           =   1020
   End
   Begin VB.Label lblCompanyProduct 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema para Análisis y Balanceo Dinámico de Rotores Rígidos y Flexibles."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   915
      Left            =   1980
      TabIndex        =   1
      Top             =   1290
      Width           =   7290
   End
   Begin VB.Label lblWarning 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmAcercade.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   885
      Left            =   90
      TabIndex        =   0
      Top             =   2850
      Width           =   9405
   End
   Begin VB.Label lblProductName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "VibraMec"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   765
      Left            =   3180
      TabIndex        =   2
      Top             =   540
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   1710
      Left            =   210
      Picture         =   "frmAcercade.frx":02E7
      Top             =   780
      Width           =   1815
   End
End
Attribute VB_Name = "frmAcercade"
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

Private Const DURACION As Long = 3

Private Sub Form_Load()
  
  Dim El_Tiempo As Long
  
  ' Muestra la ventana y le establece la animación
   'On Error GoTo Err_Proc

  Call Animar(Me, 500, AW_CENTER Or AW_ACTIVATE)
  lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
  'lblProductName.Caption = App.Title
  'lblCompanyProduct.Caption =    'App.Comments   'App.CompanyName
  'lblCopyright.Caption = App.LegalCopyright
  'lblCompany.Caption = App.LegalTrademarks
  lblPlatform.Caption = "Plataforma Windows (R)"
  'Almacena el punto de partida para hacer el retardo en segundos
  El_Tiempo = Timer + DURACION
  ' ... Espera
  Do While Timer < El_Tiempo
     DoEvents
  Loop
  ' Se cierra desde SubMain
  Listo = True

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmAcercade", "Form_Load"
   Err.Clear
   Resume Exit_Proc

End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    Unload Me
'End Sub

