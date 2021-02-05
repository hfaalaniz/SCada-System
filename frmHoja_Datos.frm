VERSION 5.00
Begin VB.Form frmHoja_Datos 
   Caption         =   "Hoja_Datos"
   ClientHeight    =   11340
   ClientLeft      =   1170
   ClientTop       =   450
   ClientWidth     =   22110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11340
   ScaleWidth      =   22110
   Begin VB.CheckBox chkFields 
      DataField       =   "Lectura_Ve"
      Enabled         =   0   'False
      Height          =   285
      Index           =   30
      Left            =   9060
      TabIndex        =   53
      Top             =   5790
      Width           =   225
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Lectura_Ve"
      Enabled         =   0   'False
      Height          =   285
      Index           =   25
      Left            =   9060
      TabIndex        =   48
      Top             =   5310
      Width           =   225
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Lectura_Ve"
      Enabled         =   0   'False
      Height          =   285
      Index           =   24
      Left            =   6870
      TabIndex        =   47
      Top             =   5310
      Width           =   225
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Lectura_Ve"
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   7890
      TabIndex        =   16
      Top             =   2250
      Width           =   225
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   22110
      TabIndex        =   76
      Top             =   10740
      Width           =   22110
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cerrar"
         Height          =   300
         Left            =   4675
         TabIndex        =   81
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Reno&var"
         Height          =   300
         Left            =   3521
         TabIndex        =   80
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   2367
         TabIndex        =   79
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Modificar"
         Height          =   300
         Left            =   1213
         TabIndex        =   78
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Nuevo"
         Height          =   300
         Left            =   59
         TabIndex        =   77
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "A&ctualizar"
         Height          =   300
         Left            =   59
         TabIndex        =   82
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   300
         Left            =   1213
         TabIndex        =   83
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   22110
      TabIndex        =   70
      Top             =   11040
      Width           =   22110
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frmHoja_Datos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmHoja_Datos.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmHoja_Datos.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmHoja_Datos.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   75
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fecha"
      Enabled         =   0   'False
      Height          =   285
      Index           =   34
      Left            =   10590
      TabIndex        =   58
      Top             =   6495
      Width           =   1000
   End
   Begin VB.TextBox txtFields 
      DataField       =   "AprobadoPor"
      Enabled         =   0   'False
      Height          =   285
      Index           =   33
      Left            =   6360
      TabIndex        =   57
      Top             =   6510
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "EjecutadoPor"
      Enabled         =   0   'False
      Height          =   285
      Index           =   32
      Left            =   1800
      TabIndex        =   55
      Top             =   6495
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano2_Radio"
      Enabled         =   0   'False
      Height          =   285
      Index           =   31
      Left            =   10710
      TabIndex        =   54
      Top             =   5835
      Width           =   900
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Plano2_G_Horario"
      Enabled         =   0   'False
      Height          =   285
      Index           =   29
      Left            =   6870
      TabIndex        =   52
      Top             =   5775
      Width           =   225
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano2_Angulo"
      Enabled         =   0   'False
      Height          =   285
      Index           =   28
      Left            =   4260
      TabIndex        =   51
      Top             =   5775
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano2_Correccion"
      Enabled         =   0   'False
      Height          =   285
      Index           =   27
      Left            =   1800
      TabIndex        =   50
      Top             =   5790
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano1_Radio"
      Enabled         =   0   'False
      Height          =   285
      Index           =   26
      Left            =   10710
      TabIndex        =   49
      Top             =   5325
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano1_Angulo"
      Enabled         =   0   'False
      Height          =   285
      Index           =   23
      Left            =   4260
      TabIndex        =   46
      Top             =   5295
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano1_Correccion"
      Enabled         =   0   'False
      Height          =   285
      Index           =   22
      Left            =   1800
      TabIndex        =   45
      Top             =   5325
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano2_a2_2"
      Enabled         =   0   'False
      Height          =   285
      Index           =   21
      Left            =   9210
      TabIndex        =   43
      Top             =   3780
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano2_V2_2"
      Enabled         =   0   'False
      Height          =   285
      Index           =   20
      Left            =   6840
      TabIndex        =   41
      Top             =   3795
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano1_a1_2"
      Enabled         =   0   'False
      Height          =   285
      Index           =   19
      Left            =   4290
      TabIndex        =   39
      Top             =   3825
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano1_V1_2"
      Enabled         =   0   'False
      Height          =   285
      Index           =   18
      Left            =   1830
      TabIndex        =   37
      Top             =   3810
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano2_a2_1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   17
      Left            =   9210
      TabIndex        =   35
      Top             =   3345
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano2_V2_1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   16
      Left            =   6840
      TabIndex        =   33
      Top             =   3345
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano1_a1_1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   15
      Left            =   4290
      TabIndex        =   31
      Top             =   3360
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano1_V1_1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   14
      Left            =   1830
      TabIndex        =   29
      Top             =   3375
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano2_a2_0"
      Enabled         =   0   'False
      Height          =   285
      Index           =   13
      Left            =   9210
      TabIndex        =   27
      Top             =   2865
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano2_V2_0"
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   6840
      TabIndex        =   25
      Top             =   2880
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano1_a1_0"
      Enabled         =   0   'False
      Height          =   285
      Index           =   11
      Left            =   4290
      TabIndex        =   23
      Top             =   2895
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Plano1_V1_0"
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   1830
      TabIndex        =   21
      Top             =   2925
      Width           =   900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Unidades"
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   9060
      TabIndex        =   19
      Top             =   2250
      Width           =   975
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Lectura_Ve"
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   6630
      TabIndex        =   15
      Top             =   2250
      Width           =   225
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Lectura_AC"
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   5370
      TabIndex        =   13
      Top             =   2220
      Width           =   195
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Radio"
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   3390
      TabIndex        =   11
      Top             =   2220
      Width           =   915
   End
   Begin VB.TextBox txtFields 
      DataField       =   "M_Prueba"
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   1830
      TabIndex        =   9
      Top             =   2235
      Width           =   885
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Descripción"
      Enabled         =   0   'False
      Height          =   675
      Index           =   3
      Left            =   1860
      TabIndex        =   7
      Top             =   1440
      Width           =   6075
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Velocidad"
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   6450
      TabIndex        =   5
      Top             =   1065
      Width           =   1485
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Rotor"
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1860
      TabIndex        =   3
      Text            =   "Descripcion del rotor"
      Top             =   1065
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ID_Hoja_Datos"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2610
      TabIndex        =   1
      Top             =   630
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   7185
      Left            =   12780
      Picture         =   "frmHoja_Datos.frx":0D08
      Stretch         =   -1  'True
      Top             =   3450
      Width           =   9030
   End
   Begin VB.Image Image1 
      Height          =   3390
      Left            =   12780
      Picture         =   "frmHoja_Datos.frx":17A65
      Stretch         =   -1  'True
      Top             =   30
      Width           =   4950
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   195
      Index           =   34
      Left            =   9990
      TabIndex        =   69
      Top             =   6495
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "AprobadoPor:"
      Height          =   195
      Index           =   33
      Left            =   5280
      TabIndex        =   68
      Top             =   6510
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EjecutadoPor:"
      Height          =   195
      Index           =   32
      Left            =   690
      TabIndex        =   67
      Top             =   6495
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano2_Radio:"
      Height          =   195
      Index           =   31
      Left            =   9555
      TabIndex        =   66
      Top             =   5835
      Width           =   1050
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano2_G_Antihorario:"
      Height          =   195
      Index           =   30
      Left            =   7365
      TabIndex        =   65
      Top             =   5790
      Width           =   1590
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano2_G_Horario:"
      Height          =   195
      Index           =   29
      Left            =   5415
      TabIndex        =   64
      Top             =   5775
      Width           =   1350
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano2_Angulo:"
      Height          =   195
      Index           =   28
      Left            =   3030
      TabIndex        =   63
      Top             =   5775
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano2_Correccion:"
      Height          =   195
      Index           =   27
      Left            =   300
      TabIndex        =   62
      Top             =   5790
      Width           =   1395
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano1_Radio:"
      Height          =   195
      Index           =   26
      Left            =   9555
      TabIndex        =   61
      Top             =   5325
      Width           =   1050
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano1_G_Antihorario:"
      Height          =   195
      Index           =   25
      Left            =   7365
      TabIndex        =   60
      Top             =   5325
      Width           =   1590
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano1_G_Horario:"
      Height          =   195
      Index           =   24
      Left            =   5415
      TabIndex        =   59
      Top             =   5340
      Width           =   1350
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano1_Angulo:"
      Height          =   195
      Index           =   23
      Left            =   3030
      TabIndex        =   56
      Top             =   5295
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano1_Correccion:"
      Height          =   195
      Index           =   22
      Left            =   300
      TabIndex        =   44
      Top             =   5325
      Width           =   1395
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano2_a2_2:"
      Height          =   195
      Index           =   21
      Left            =   8115
      TabIndex        =   42
      Top             =   3780
      Width           =   990
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano2_V2_2:"
      Height          =   195
      Index           =   20
      Left            =   5730
      TabIndex        =   40
      Top             =   3795
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano1_a1_2:"
      Height          =   195
      Index           =   19
      Left            =   3195
      TabIndex        =   38
      Top             =   3825
      Width           =   990
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano1_V1_2:"
      Height          =   195
      Index           =   18
      Left            =   720
      TabIndex        =   36
      Top             =   3810
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano2_a2_1:"
      Height          =   195
      Index           =   17
      Left            =   8115
      TabIndex        =   34
      Top             =   3345
      Width           =   990
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano2_V2_1:"
      Height          =   195
      Index           =   16
      Left            =   5730
      TabIndex        =   32
      Top             =   3345
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano1_a1_1:"
      Height          =   195
      Index           =   15
      Left            =   3195
      TabIndex        =   30
      Top             =   3360
      Width           =   990
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano1_V1_1:"
      Height          =   195
      Index           =   14
      Left            =   720
      TabIndex        =   28
      Top             =   3375
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano2_a2_0:"
      Height          =   195
      Index           =   13
      Left            =   8115
      TabIndex        =   26
      Top             =   2865
      Width           =   990
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano2_V2_0:"
      Height          =   195
      Index           =   12
      Left            =   5730
      TabIndex        =   24
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano1_a1_0:"
      Height          =   195
      Index           =   11
      Left            =   3195
      TabIndex        =   22
      Top             =   2895
      Width           =   990
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Plano1_V1_0:"
      Height          =   195
      Index           =   10
      Left            =   720
      TabIndex        =   20
      Top             =   2925
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Unidades:"
      Height          =   195
      Index           =   9
      Left            =   8265
      TabIndex        =   18
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Lectura_De:"
      Height          =   195
      Index           =   8
      Left            =   6930
      TabIndex        =   17
      Top             =   2280
      Width           =   885
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Lectura_Ve:"
      Height          =   195
      Index           =   7
      Left            =   5655
      TabIndex        =   14
      Top             =   2265
      Width           =   870
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Lectura_AC:"
      Height          =   195
      Index           =   6
      Left            =   4380
      TabIndex        =   12
      Top             =   2220
      Width           =   885
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Radio:"
      Height          =   195
      Index           =   5
      Left            =   2820
      TabIndex        =   10
      Top             =   2235
      Width           =   465
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "M_Prueba:"
      Height          =   195
      Index           =   4
      Left            =   945
      TabIndex        =   8
      Top             =   2235
      Width           =   780
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      Height          =   195
      Index           =   3
      Left            =   870
      TabIndex        =   6
      Top             =   1440
      Width           =   885
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Velocidad:"
      Height          =   195
      Index           =   2
      Left            =   5595
      TabIndex        =   4
      Top             =   1065
      Width           =   750
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Rotor:"
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   1065
      Width           =   435
   End
   Begin VB.Label lblLabels 
      Caption         =   "ID_Hoja_Datos:"
      Height          =   255
      Index           =   0
      Left            =   690
      TabIndex        =   0
      Top             =   630
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmHoja_Datos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Form_Load()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VibraMec\VIBRAMEC.mdb;"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select ID_Hoja_Datos,Rotor,Velocidad,Descripción,M_Prueba,Radio,Lectura_AC,Lectura_Ve,Lectura_De,Unidades,Plano1_V1_0,Plano1_a1_0,Plano2_V2_0,Plano2_a2_0,Plano1_V1_1,Plano1_a1_1,Plano2_V2_1,Plano2_a2_1,Plano1_V1_2,Plano1_a1_2,Plano2_V2_2,Plano2_a2_2,Plano1_Correccion,Plano1_Angulo,Plano1_G_Horario,Plano1_G_Antihorario,Plano1_Radio,Plano2_Correccion,Plano2_Angulo,Plano2_G_Horario,Plano2_G_Antihorario,Plano2_Radio,EjecutadoPor,AprobadoPor,Fecha from Hoja_Datos Order by Fecha", db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Enlaza los cuadros de texto con el proveedor de datos
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  Dim oCheck As CheckBox
  'Enlaza las casillas de verificación con el proveedor de datos
  For Each oCheck In Me.chkFields
    Set oCheck.DataSource = adoPrimaryRS
  Next

  mbDataChanged = False
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  lblStatus.Caption = "Hoja: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    lblStatus.Caption = "Ingresar nueva hoja de datos"
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  adoPrimaryRS.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Modificar hoja de datos"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'va al nuevo registro
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'ha sobrepasado el final; vuelva atrás
    adoPrimaryRS.MoveLast
  End If
  'muestra el registro actual
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'ha sobrepasado el final; vuelva atrás
    adoPrimaryRS.MoveFirst
  End If
  'muestra el registro actual
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

