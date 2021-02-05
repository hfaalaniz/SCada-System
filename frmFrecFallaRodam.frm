VERSION 5.00
Begin VB.Form frmFrecFallaRodam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C·lculo de la Frecuencia de falla de Rodamientos"
   ClientHeight    =   11115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   19005
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   75
      Left            =   150
      TabIndex        =   23
      Top             =   10440
      Width           =   15675
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frecuencias de Falla en Hz."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3555
      Left            =   315
      TabIndex        =   10
      Top             =   4455
      Width           =   6915
      Begin VB.TextBox txtBPFO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtBPFI 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtBPF 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   12
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtFTF 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblBPFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frecuencia de defecto en la pista-cubeta exterior en Hz."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   27
         Top             =   480
         Width           =   4065
      End
      Begin VB.Label lblBPFI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frecuencia de defecto en la pista-cubeta interior en Hz."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   26
         Top             =   1200
         Width           =   4005
      End
      Begin VB.Label lbl_BPF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frecuencia de defecto en el elemento rodante-bola en Hz."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   25
         Top             =   1920
         Width           =   4185
      End
      Begin VB.Label lblFTF 
         BackStyle       =   0  'Transparent
         Caption         =   "Frecuencia de rotaciÛn del porta elemento o jaula que contiene los elementos rodantes."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   24
         Top             =   2640
         Width           =   6360
      End
   End
   Begin VB.Frame frmIngDatos 
      Caption         =   "Ingreso de Datos para el c·lculo."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Left            =   300
      TabIndex        =   9
      Top             =   180
      Width           =   7635
      Begin VB.ComboBox cmbBuscarRodam 
         Height          =   315
         Left            =   420
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   720
         Width           =   6735
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   360
         Left            =   4920
         TabIndex        =   6
         ToolTipText     =   "Buscar rodamiento por codigo y marca."
         Top             =   3600
         Width           =   990
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "Calcular"
         Height          =   360
         Left            =   3840
         TabIndex        =   5
         ToolTipText     =   "Iniciar c·lculo..."
         Top             =   3600
         Width           =   990
      End
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   360
         Left            =   6240
         TabIndex        =   7
         ToolTipText     =   "Limpiar ˙ltimo c·lculo."
         Top             =   3600
         Width           =   990
      End
      Begin VB.TextBox txtRPM 
         Height          =   375
         Left            =   5220
         TabIndex        =   1
         ToolTipText     =   "Velocidad del aro interior en RPM."
         Top             =   1260
         Width           =   1335
      End
      Begin VB.TextBox txtDiamBolas 
         Height          =   375
         Left            =   5220
         TabIndex        =   2
         ToolTipText     =   "Di·metro del elemento rodante en mm."
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtDiamCenBolas 
         Height          =   375
         Left            =   5220
         TabIndex        =   3
         ToolTipText     =   "di·metro efectivo (di·metro entre los centros de los elementos rodantes) en mm."
         Top             =   2340
         Width           =   1335
      End
      Begin VB.TextBox txtAngCtto 
         Height          =   375
         Left            =   5220
         TabIndex        =   4
         ToolTipText     =   "¡ngulo de contacto en grados."
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblSelRodCodMarca 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccionar rodamiento, codigo y marca..."
         Height          =   195
         Left            =   420
         TabIndex        =   28
         Top             =   480
         Width           =   3030
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "grados."
         Height          =   195
         Left            =   6660
         TabIndex        =   22
         Top             =   2940
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm."
         Height          =   195
         Left            =   6660
         TabIndex        =   21
         Top             =   2400
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm."
         Height          =   195
         Left            =   6660
         TabIndex        =   20
         Top             =   1920
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RPM."
         Height          =   195
         Left            =   6660
         TabIndex        =   19
         Top             =   1380
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¡ngulo de contacto en grados."
         Height          =   195
         Left            =   2940
         TabIndex        =   18
         Top             =   2940
         Width           =   2220
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Di·metro efectivo entre los centros de los elementos rodantes en mm."
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2460
         Width           =   5040
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Di·metro del elemento rodante en mm."
         Height          =   195
         Left            =   2400
         TabIndex        =   16
         Top             =   1920
         Width           =   2790
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidad del aro interior en RPM."
         Height          =   195
         Left            =   2760
         TabIndex        =   15
         Top             =   1380
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   11205
      TabIndex        =   8
      Top             =   10665
      Width           =   990
   End
   Begin VB.Label Label12 
      Caption         =   $"frmFrecFallaRodam.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8055
      TabIndex        =   32
      Top             =   2565
      Width           =   7935
   End
   Begin VB.Label Label11 
      Caption         =   $"frmFrecFallaRodam.frx":00B8
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8055
      TabIndex        =   31
      Top             =   1800
      Width           =   7935
   End
   Begin VB.Label Label10 
      Caption         =   $"frmFrecFallaRodam.frx":0168
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8055
      TabIndex        =   30
      Top             =   1035
      Width           =   7935
   End
   Begin VB.Label Label9 
      Caption         =   $"frmFrecFallaRodam.frx":0227
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8055
      TabIndex        =   29
      Top             =   270
      Width           =   7935
   End
   Begin VB.Image imgRodam 
      Height          =   4860
      Left            =   7695
      Picture         =   "frmFrecFallaRodam.frx":02E6
      Top             =   4635
      Width           =   4050
   End
End
Attribute VB_Name = "frmFrecFallaRodam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmFrecFallaRodam
'    Project    : prjFrecRodamientos
'
'    Description: [C·culo de la frecuencia de falla de rodamientos.]
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

Dim v_BPFO, v_BPFI, v_BPF, v_FTF As Double

'txtRPM             - velocidad de la pista interior.
'txtDiamBolas       - di·metro de las bolas.
'txtDiamCenBolas    - di·metro entre centros de bolas.
'txtAngCtto         - ·ngulo de contacto.
'txtBPFO            - frecuencia de defecto en la pista exterior.
'txtBPFI            - frecuencia de defecto en la pista interior.
'txtBPF             - frecuencia de defecto en el elemento rodante.
'txtFTF             - frecuencia de defecto en la jaula.
'

Private Sub cmdCalcular_Click()
'Las frecuencias de fallo estan relacionadas a la velocidad de rotacion del anillo interno ÅgNiÅh (ya
'que el anillo interior y el eje tienen la misma velocidad de rotacion, de aqui en adelante Ni sera
'simplemente N), el diametro de inclinacion del rodamiento ÅgDÅh, el diametro del elemento
'rodante ÅgdÅh, el numero de bolas o de rodillos ÅgnÅh, y el angulo de contacto ÅgÉ∆Åh.
End Sub
