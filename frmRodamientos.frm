VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRodamientos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccione Rodamiento para iniciar el cálculo..."
   ClientHeight    =   11040
   ClientLeft      =   1095
   ClientTop       =   375
   ClientWidth     =   19320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11040
   ScaleWidth      =   19320
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox iXYPlotX1 
      Height          =   3075
      Left            =   7785
      ScaleHeight     =   3015
      ScaleWidth      =   11430
      TabIndex        =   73
      Top             =   3510
      Width           =   11490
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
      Left            =   7785
      TabIndex        =   64
      Top             =   6705
      Width           =   6915
      Begin VB.TextBox txtFTF 
         Alignment       =   2  'Center
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
         TabIndex        =   68
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtBPF 
         Alignment       =   2  'Center
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
         TabIndex        =   67
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtBPFI 
         Alignment       =   2  'Center
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
         TabIndex        =   66
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtBPFO 
         Alignment       =   2  'Center
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
         TabIndex        =   65
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblFTF 
         BackStyle       =   0  'Transparent
         Caption         =   "Frecuencia de rotación del porta elemento o jaula que contiene los elementos rodantes."
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
         TabIndex        =   72
         Top             =   2640
         Width           =   6360
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
         TabIndex        =   71
         Top             =   1920
         Width           =   4185
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
         TabIndex        =   70
         Top             =   1200
         Width           =   4005
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
         Left            =   1665
         TabIndex        =   69
         Top             =   450
         Width           =   4065
      End
   End
   Begin VB.Frame frmIngDatos 
      Caption         =   "Ingreso de Datos para el cálculo."
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
      Left            =   90
      TabIndex        =   46
      Top             =   5760
      Width           =   7635
      Begin VB.TextBox txtAngCtto 
         Height          =   375
         Left            =   5220
         TabIndex        =   54
         ToolTipText     =   "Ángulo de contacto en grados."
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtDiamCenBolas 
         Height          =   375
         Left            =   5220
         TabIndex        =   53
         ToolTipText     =   "diámetro efectivo (diámetro entre los centros de los elementos rodantes) en mm."
         Top             =   2340
         Width           =   1335
      End
      Begin VB.TextBox txtDiamBolas 
         Height          =   375
         Left            =   5220
         TabIndex        =   52
         ToolTipText     =   "Diámetro del elemento rodante en mm."
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtRPM 
         Height          =   375
         Left            =   5220
         TabIndex        =   51
         ToolTipText     =   "Velocidad del aro interior en RPM."
         Top             =   1260
         Width           =   1335
      End
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   360
         Left            =   6240
         TabIndex        =   50
         ToolTipText     =   "Limpiar último cálculo."
         Top             =   3600
         Width           =   990
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "Calcular"
         Height          =   360
         Left            =   3840
         TabIndex        =   49
         ToolTipText     =   "Iniciar cálculo..."
         Top             =   3600
         Width           =   990
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   360
         Left            =   4920
         TabIndex        =   48
         ToolTipText     =   "Buscar rodamiento por codigo y marca."
         Top             =   3600
         Width           =   990
      End
      Begin VB.ComboBox cmbBuscarRodam 
         Height          =   315
         Left            =   420
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   720
         Width           =   6735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidad del aro interior en RPM."
         Height          =   195
         Left            =   2760
         TabIndex        =   63
         Top             =   1380
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diámetro del elemento rodante en mm."
         Height          =   195
         Left            =   2400
         TabIndex        =   62
         Top             =   1920
         Width           =   2790
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diámetro efectivo entre los centros de los elementos rodantes en mm."
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   2460
         Width           =   5040
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ángulo de contacto en grados."
         Height          =   195
         Left            =   2940
         TabIndex        =   60
         Top             =   2940
         Width           =   2220
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RPM."
         Height          =   195
         Left            =   6660
         TabIndex        =   59
         Top             =   1380
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm."
         Height          =   195
         Left            =   6660
         TabIndex        =   58
         Top             =   1920
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm."
         Height          =   195
         Left            =   6660
         TabIndex        =   57
         Top             =   2400
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "grados."
         Height          =   195
         Left            =   6660
         TabIndex        =   56
         Top             =   2940
         Width           =   555
      End
      Begin VB.Label lblSelRodCodMarca 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccionar rodamiento, codigo y marca..."
         Height          =   195
         Left            =   420
         TabIndex        =   55
         Top             =   480
         Width           =   3030
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "AD1"
      Height          =   285
      Index           =   27
      Left            =   2055
      TabIndex        =   44
      Top             =   5415
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Masa"
      Height          =   285
      Index           =   26
      Left            =   2055
      TabIndex        =   42
      Top             =   5100
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "RPMlim"
      Height          =   285
      Index           =   25
      Left            =   2055
      TabIndex        =   40
      Top             =   4785
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "RPMref"
      Height          =   285
      Index           =   24
      Left            =   2055
      TabIndex        =   38
      Top             =   4455
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Pu"
      Height          =   285
      Index           =   23
      Left            =   2055
      TabIndex        =   36
      Top             =   4140
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Co"
      Height          =   285
      Index           =   22
      Left            =   2055
      TabIndex        =   34
      Top             =   3825
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1"
      Height          =   285
      Index           =   21
      Left            =   2055
      TabIndex        =   32
      Top             =   3495
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C"
      Height          =   285
      Index           =   20
      Left            =   2055
      TabIndex        =   30
      Top             =   3180
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "B"
      Height          =   285
      Index           =   19
      Left            =   2055
      TabIndex        =   28
      Top             =   2865
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "D_Exterior"
      Height          =   285
      Index           =   18
      Left            =   2055
      TabIndex        =   26
      Top             =   2535
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "d_Interior"
      Height          =   285
      Index           =   17
      Left            =   2055
      TabIndex        =   24
      Top             =   2220
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Nombre"
      Height          =   285
      Index           =   16
      Left            =   2055
      TabIndex        =   22
      Top             =   1905
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Marca"
      Height          =   285
      Index           =   15
      Left            =   2055
      TabIndex        =   20
      Top             =   1575
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "IDRodam"
      Height          =   285
      Index           =   14
      Left            =   2055
      TabIndex        =   18
      Top             =   1260
      Width           =   3375
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5580
      Top             =   1980
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRodamientos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRodamientos.frx":0FD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRodamientos.frx":12EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRodamientos.frx":1606
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRodamientos.frx":1920
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   19320
      _ExtentX        =   34078
      _ExtentY        =   1482
      ButtonWidth     =   1773
      ButtonHeight    =   1429
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Abrir"
            Key             =   "abrir"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Guardar"
            Key             =   "guardar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleccionar"
            Key             =   "combo1"
            Style           =   4
            Object.Width           =   2000
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar"
            Key             =   "cerrar"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   420
         Width           =   1935
      End
   End
   Begin VB.ComboBox cmbBuscar 
      Height          =   315
      Left            =   135
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   900
      Width           =   4320
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   19320
      TabIndex        =   6
      Top             =   10440
      Width           =   19320
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   300
         Left            =   1213
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "A&ctualizar"
         Height          =   300
         Left            =   59
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cerrar"
         Height          =   300
         Left            =   4675
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Reno&var"
         Height          =   300
         Left            =   3521
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   2367
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edición"
         Height          =   300
         Left            =   1213
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Ag&regar"
         Height          =   300
         Left            =   59
         TabIndex        =   7
         Top             =   0
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
      ScaleWidth      =   19320
      TabIndex        =   0
      Top             =   10740
      Width           =   19320
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frmRodamientos.frx":1C3A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmRodamientos.frx":1F7C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmRodamientos.frx":22BE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmRodamientos.frx":2600
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   5
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.Label lblCount 
      Caption         =   "lblCount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4545
      TabIndex        =   45
      Top             =   855
      Width           =   1545
   End
   Begin VB.Label lblLabels 
      Caption         =   "AD1:"
      Height          =   255
      Index           =   27
      Left            =   135
      TabIndex        =   43
      Top             =   5415
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Masa:"
      Height          =   255
      Index           =   26
      Left            =   135
      TabIndex        =   41
      Top             =   5100
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "RPMlim:"
      Height          =   255
      Index           =   25
      Left            =   135
      TabIndex        =   39
      Top             =   4785
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "RPMref:"
      Height          =   255
      Index           =   24
      Left            =   135
      TabIndex        =   37
      Top             =   4455
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Pu:"
      Height          =   255
      Index           =   23
      Left            =   135
      TabIndex        =   35
      Top             =   4140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Co:"
      Height          =   255
      Index           =   22
      Left            =   135
      TabIndex        =   33
      Top             =   3825
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "C1:"
      Height          =   255
      Index           =   21
      Left            =   135
      TabIndex        =   31
      Top             =   3495
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "C:"
      Height          =   255
      Index           =   20
      Left            =   135
      TabIndex        =   29
      Top             =   3180
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "B:"
      Height          =   255
      Index           =   19
      Left            =   135
      TabIndex        =   27
      Top             =   2865
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "D_Exterior:"
      Height          =   255
      Index           =   18
      Left            =   135
      TabIndex        =   25
      Top             =   2535
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "d_Interior:"
      Height          =   255
      Index           =   17
      Left            =   135
      TabIndex        =   23
      Top             =   2220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nombre:"
      Height          =   255
      Index           =   16
      Left            =   135
      TabIndex        =   21
      Top             =   1905
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Marca:"
      Height          =   255
      Index           =   15
      Left            =   135
      TabIndex        =   19
      Top             =   1575
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "IDRodam:"
      Height          =   255
      Index           =   2
      Left            =   135
      TabIndex        =   17
      Top             =   1260
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   2370
      Left            =   6525
      Picture         =   "frmRodamientos.frx":2942
      Stretch         =   -1  'True
      Top             =   990
      Width           =   12165
   End
End
Attribute VB_Name = "frmRodamientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As connection
Dim WithEvents adoRodam As Recordset
Attribute adoRodam.VB_VarHelpID = -1
Dim WithEvents adoRodam_1 As Recordset
Attribute adoRodam_1.VB_VarHelpID = -1
Dim WithEvents adoRodamClone As Recordset
Attribute adoRodamClone.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Dim v_BPFO, v_BPFI, v_BPF, v_FTF As Double, n As Integer

'txtRPM             - velocidad de la pista interior.
'txtDiamBolas       - diámetro de las bolas.
'txtDiamCenBolas    - diámetro entre centros de bolas.
'txtAngCtto         - ángulo de contacto.
'txtBPFO            - frecuencia de defecto en la pista exterior.
'txtBPFI            - frecuencia de defecto en la pista interior.
'txtBPF             - frecuencia de defecto en el elemento rodante.
'txtFTF             - frecuencia de defecto en la jaula.
'
Private Sub cmdCalcular_Click()

n = 1450
'Las frecuencias de fallo estan relacionadas a la velocidad de rotacion del anillo interno gNih (ya
'que el anillo interior y el eje tienen la misma velocidad de rotacion, de aqui en adelante Ni sera
'simplemente N), el diametro de inclinacion del rodamiento gDh, el diametro del elemento
'rodante gdh, el numero de bolas o de rodillos gnh, y el angulo de contacto gƒÆh.
'd_Interior = txtFields(17)  -----  d_Exterior = txtFields(18)
v_BPF = 0.5 * n * (Val(txtFields(18).Text) / (Val(txtFields(17).Text)) * (1 - (Val(txtFields(17).Text) / (Val(txtFields(18).Text))) ^ 2))
txtBPF.Text = Format(v_BPF, "###0.00") & " Hz"

v_FTF = 0.5 * n * (1 - (Val(txtFields(17).Text) / (Val(txtFields(18).Text))))
txtFTF.Text = Format(v_FTF, "###0.00") & " Hz"

v_BPFI = 0.5 * n * (1 + (Val(txtFields(17).Text) / (Val(txtFields(18).Text))))
txtBPFI = Format(v_BPFI, "###0.00") & " Hz"

v_BPFO = 0.5 * n * (1 - (Val(txtFields(17).Text) / (Val(txtFields(18).Text))))
txtBPFO = Format(v_BPFO, "###0.00") & " Hz"
End Sub

Private Sub cmbBuscar_Change()
    Call Cargar_Combo
End Sub

Private Sub cmbBuscar_Click()
'adoRodam_1.Close
'Set adoRodam = New Recordset
'adoRodam_1.Close
'adoRodam_1.Open "select * from tblRodam WHERE Nombre='" & cmbBuscar.Text & "'", db, adOpenStatic, adLockOptimistic
'  Dim oText As TextBox
'  'Enlaza los cuadros de texto con el proveedor de datos
'  For Each oText In Me.txtFields
'    Set oText.DataSource = adoRodam_1
'   Next
'adoRodam_1.Close
'adoRodam_1.Open "select * from tblRodam", db, adOpenStatic, adLockOptimistic

End Sub

Sub CrearToolBar()
   ' Configura el control ComboBox para colocarlo en
    ' la misma posición que el objeto Button con el
   ' estilo Placeholder (key = "combo1").
   With Combo1
      .Width = Toolbar1.Buttons("combo1").Width
      .Top = Toolbar1.Buttons("combo1").Top + 2000
      .Left = Toolbar1.Buttons("combo1").Left
      .AddItem "Sistema" ' Agrega colores al texto.
      '.AddItem "Azul"
      '.AddItem "Rojo"
      '.ListIndex = 0
   End With
End Sub

Sub Cargar_Combo()
Dim intI As Integer
Set adoRodam_1 = New Recordset
adoRodam_1.Open "select * from tblRodam Order by Nombre", db, adOpenStatic, adLockOptimistic
adoRodam_1.MoveFirst
    If adoRodam_1.EOF Then Exit Sub
    intI = 1
    With adoRodam_1
        Do Until .EOF
            cmbBuscar.AddItem adoRodam_1!Nombre     'Nombre
            .MoveNext
            intI = intI + 1
        Loop
        .Close
    End With
cmbBuscar.ListIndex = 0
End Sub

Private Sub Form_Load()

  Set db = New connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\VIBRAMEC.mdb;"

  Set adoRodam = New Recordset
  adoRodam.Open "select * from tblRodam Order by Nombre", db, adOpenStatic, adLockOptimistic
  lblCount.Caption = adoRodam.RecordCount
Set adoRodamClone = adoRodam.Clone
adoRodam.MoveFirst

  Dim oText As TextBox
  'Enlaza los cuadros de texto con el proveedor de datos
  For Each oText In Me.txtFields
    Set oText.DataSource = adoRodam
  Next

Call CrearToolBar
Call Cargar_Combo
mbDataChanged = False
cmdCalcular_Click
End Sub

Private Sub Form_Resize()
    On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
   ' Configura el control ComboBox.
   With Combo1
      .Width = Toolbar1.Buttons("combo1").Width
      .Top = Toolbar1.Buttons("combo1").Top
      .Left = Toolbar1.Buttons("combo1").Left
   End With
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

Private Sub adoRodam_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  lblStatus.Caption = "Record: " & CStr(adoRodam.AbsolutePosition)
End Sub

Private Sub adoRodam_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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
  With adoRodam
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    lblStatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With adoRodam
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
  adoRodam.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Modificar registro"
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
  adoRodam.CancelUpdate
  If mvBookMark > 0 Then
    adoRodam.Bookmark = mvBookMark
  Else
    adoRodam.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  adoRodam.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoRodam.MoveLast              'va al nuevo registro
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

  adoRodam.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoRodam.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not adoRodam.EOF Then adoRodam.MoveNext
  If adoRodam.EOF And adoRodam.RecordCount > 0 Then
    Beep
     'ha sobrepasado el final; vuelva atrás
    adoRodam.MoveLast
  End If
  'muestra el registro actual
  mbDataChanged = False
  cmdCalcular_Click
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not adoRodam.BOF Then adoRodam.MovePrevious
  If adoRodam.BOF And adoRodam.RecordCount > 0 Then
    Beep
    'ha sobrepasado el final; vuelva atrás
    adoRodam.MoveFirst
  End If
  'muestra el registro actual
  mbDataChanged = False
  cmdCalcular_Click
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

Private Sub Toolbar1_ButtonClick(ByVal Button As Button)
   ' Utiliza la propiedad Key con la instrucción
   ' SelectCase para especificar una acción.
   Select Case Button.Key
   Case Is = "abrir"         ' Abre archivo.
      'MsgBox "Agregue código para abrir el archivo"
      frmFrecFallaRodam.Show vbModal
   Case Is = "guardar"           ' Guarda archivo.
      MsgBox "Agregue código para guardar el código"
   Case Is = "cerrar"
      End
   End Select
End Sub

Private Sub Combo1_Click()
   ' Cambia el color del fondo utilizando el ComboBox.
   Select Case Combo1.ListIndex
   Case 0
      Me.BackColor = &H8000000F
   Case 1
      Me.BackColor = vbBlue
   Case 2
      Me.BackColor = vbRed
   End Select
End Sub
