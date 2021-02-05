VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Begin VB.Form frmPrincipal 
   BackColor       =   &H00404040&
   Caption         =   "VibraMeK - Análisis y Balanceo de Vibraciones Mecánicas. (c) 2012 por Hugo Fabián Alaníz"
   ClientHeight    =   10635
   ClientLeft      =   3960
   ClientTop       =   -750
   ClientWidth     =   19020
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10635
   ScaleWidth      =   19020
   WindowState     =   2  'Maximized
   Begin VibraMec.ThemedComboBox ThemedComboBox1 
      Left            =   5970
      Top             =   630
      _ExtentX        =   556
      _ExtentY        =   529
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   7
      Top             =   10170
      Width           =   19020
      _ExtentX        =   33549
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Leer_Escribir 
      Interval        =   1
      Left            =   5940
      Top             =   1140
   End
   Begin VB.PictureBox picT 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   18960
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   19020
      Begin MSComctlLib.ListView ListView1 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   0
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   2100
      Left            =   3300
      ScaleHeight     =   914.43
      ScaleMode       =   0  'User
      ScaleWidth      =   624
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox picR 
      BackColor       =   &H00505050&
      Height          =   1980
      Left            =   3360
      ScaleHeight     =   1920
      ScaleWidth      =   2145
      TabIndex        =   1
      Top             =   510
      Width           =   2205
   End
   Begin VB.PictureBox picL 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   11070
      Left            =   0
      ScaleHeight     =   11040
      ScaleWidth      =   3225
      TabIndex        =   0
      Top             =   420
      Width           =   3255
      Begin TabDlg.SSTab tabMain 
         Height          =   11220
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   19791
         _Version        =   393216
         Style           =   1
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Base de Datos"
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "trvListaEmpr"
         Tab(0).Control(1)=   "XTab1"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Balanceo"
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "frPlano2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frPlano1"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Análisis"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SSTab3"
         Tab(2).Control(1)=   "frAnalisis(1)"
         Tab(2).ControlCount=   2
         Begin VB.Frame frPlano1 
            BackColor       =   &H00505050&
            Height          =   3525
            Left            =   90
            TabIndex        =   50
            Top             =   450
            Width           =   3015
            Begin VB.ComboBox cmbPlano1 
               Height          =   315
               Index           =   1
               Left            =   60
               TabIndex        =   51
               Top             =   630
               Width           =   2955
            End
            Begin ComCtl3.CoolBar CoolBar1 
               Height          =   390
               Left            =   60
               TabIndex        =   52
               Top             =   1110
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   688
               BandCount       =   2
               EmbossPicture   =   -1  'True
               _CBWidth        =   2895
               _CBHeight       =   390
               _Version        =   "6.7.9816"
               Child1          =   "Combo1"
               MinHeight1      =   315
               Width1          =   1395
               NewRow1         =   0   'False
               Child2          =   "tlbPlano1"
               MinHeight2      =   330
               Width2          =   2430
               NewRow2         =   0   'False
               Begin MSComctlLib.Toolbar Toolbar1 
                  Height          =   660
                  Left            =   1530
                  TabIndex        =   54
                  Top             =   45
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   1164
                  ButtonWidth     =   609
                  ButtonHeight    =   582
                  Style           =   1
                  ImageList       =   "ImageList1"
                  _Version        =   393216
                  BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                     NumButtons      =   4
                     BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "nuevoBal1"
                        ImageIndex      =   3
                     EndProperty
                     BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        ImageIndex      =   1
                     EndProperty
                     BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        ImageIndex      =   2
                     EndProperty
                     BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     EndProperty
                  EndProperty
               End
               Begin VB.ComboBox Combo1 
                  Height          =   315
                  Index           =   1
                  Left            =   165
                  TabIndex        =   53
                  Text            =   "Combo1"
                  Top             =   30
                  Width           =   1200
               End
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Label1"
               Height          =   195
               Left            =   1620
               TabIndex        =   87
               Top             =   1860
               Width           =   480
            End
            Begin VB.Label lblPlano 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               Caption         =   "Plano 1"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFC0&
               Height          =   405
               Index           =   1
               Left            =   0
               TabIndex        =   61
               Top             =   90
               Width           =   3015
            End
            Begin VB.Label lblFiltroP1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Filtro:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000003&
               Height          =   375
               Index           =   1
               Left            =   90
               TabIndex        =   60
               Top             =   1650
               Width           =   780
            End
            Begin VB.Label lblFaseP1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Fase:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000003&
               Height          =   375
               Index           =   1
               Left            =   90
               TabIndex        =   59
               Top             =   2010
               Width           =   645
            End
            Begin VB.Label lblRMSP1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "RMS:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   58
               Top             =   2520
               Width           =   390
            End
            Begin VB.Label lblMaxP1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Max:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   57
               Top             =   2760
               Width           =   360
            End
            Begin VB.Label lblFrecMaxP1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Frec. del Max:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   56
               Top             =   2970
               Width           =   1050
            End
            Begin VB.Label lblTacometroP1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tacómetro:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   55
               Top             =   3180
               Width           =   840
            End
            Begin VB.Shape shpPlano1 
               BorderColor     =   &H00FF8080&
               FillColor       =   &H00505050&
               FillStyle       =   0  'Solid
               Height          =   1065
               Index           =   1
               Left            =   30
               Top             =   2400
               Width           =   2955
            End
         End
         Begin VB.Frame frPlano2 
            BackColor       =   &H00505050&
            Height          =   3525
            Left            =   90
            TabIndex        =   38
            Top             =   4020
            Width           =   3015
            Begin VB.ComboBox Combo2 
               Height          =   315
               Index           =   1
               Left            =   60
               TabIndex        =   39
               Top             =   630
               Width           =   2955
            End
            Begin ComCtl3.CoolBar clbPlano2 
               Height          =   390
               Left            =   0
               TabIndex        =   40
               Top             =   1215
               Width           =   2925
               _ExtentX        =   5159
               _ExtentY        =   688
               BandCount       =   2
               ForeColor       =   5263440
               _CBWidth        =   2925
               _CBHeight       =   390
               _Version        =   "6.7.9816"
               Child1          =   "Combo1"
               MinHeight1      =   315
               Width1          =   1395
               NewRow1         =   0   'False
               Child2          =   "tlbPlano1"
               MinHeight2      =   330
               Width2          =   2430
               NewRow2         =   0   'False
               Begin VB.ComboBox Combo3 
                  Height          =   315
                  Left            =   165
                  TabIndex        =   42
                  Text            =   "Combo1"
                  Top             =   30
                  Width           =   1320
               End
               Begin MSComctlLib.Toolbar tlbPlano2 
                  Height          =   660
                  Left            =   1560
                  TabIndex        =   41
                  Top             =   30
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   1164
                  ButtonWidth     =   609
                  ButtonHeight    =   582
                  Style           =   1
                  ImageList       =   "ImageList1"
                  _Version        =   393216
                  BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                     NumButtons      =   4
                     BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        ImageIndex      =   3
                     EndProperty
                     BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        ImageIndex      =   1
                     EndProperty
                     BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        ImageIndex      =   2
                     EndProperty
                     BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     EndProperty
                  EndProperty
               End
            End
            Begin VB.Label lblPlano 
               Alignment       =   2  'Center
               BackColor       =   &H00FF0000&
               Caption         =   "Plano 2"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   405
               Index           =   2
               Left            =   5
               TabIndex        =   49
               Top             =   90
               Width           =   3015
            End
            Begin VB.Label lblTacometroP2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tacómetro:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   48
               Top             =   3180
               Width           =   840
            End
            Begin VB.Label lblFrecMaxP2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Frec. del Max:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   47
               Top             =   2970
               Width           =   1050
            End
            Begin VB.Label lblMaxP2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Max:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   46
               Top             =   2760
               Width           =   360
            End
            Begin VB.Label lblRMSP2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "RMS:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   45
               Top             =   2520
               Width           =   390
            End
            Begin VB.Label lblFaseP2 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Fase:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   90
               TabIndex        =   44
               Top             =   2010
               Width           =   645
            End
            Begin VB.Label lblFiltroP2 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Filtro:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   90
               TabIndex        =   43
               Top             =   1650
               Width           =   780
            End
            Begin VB.Shape shpPlano2 
               BorderColor     =   &H00FF8080&
               FillColor       =   &H00505050&
               FillStyle       =   0  'Solid
               Height          =   1065
               Index           =   1
               Left            =   30
               Top             =   2400
               Width           =   2955
            End
         End
         Begin VB.Frame frAnalisis 
            BackColor       =   &H00FFC0C0&
            Height          =   2925
            Index           =   1
            Left            =   -74910
            TabIndex        =   30
            Top             =   495
            Width           =   3045
            Begin VB.ComboBox cmdAnalisis1 
               Height          =   315
               Index           =   1
               Left            =   60
               TabIndex        =   32
               Top             =   630
               Width           =   2955
            End
            Begin VB.ComboBox cmdAnalisis2 
               Height          =   315
               Index           =   1
               Left            =   60
               TabIndex        =   31
               Top             =   990
               Width           =   2955
            End
            Begin MSComctlLib.Toolbar tlbAnalisis 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   33
               Top             =   1380
               Width           =   2865
               _ExtentX        =   5054
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   5
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Style           =   3
                  EndProperty
                  BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            End
            Begin VB.Label lblAnalisis 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               Caption         =   "Análisis"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   1
               Left            =   0
               TabIndex        =   37
               Top             =   90
               Width           =   3015
            End
            Begin VB.Label lblRMS 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "RMS:"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   36
               Top             =   2040
               Width           =   405
            End
            Begin VB.Label lblMax 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Max:"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   35
               Top             =   2280
               Width           =   345
            End
            Begin VB.Label lblFrecMax 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Frec. del Max:"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   34
               Top             =   2520
               Width           =   1005
            End
            Begin VB.Shape shpAnalisis 
               BorderColor     =   &H00FF8080&
               FillColor       =   &H80000002&
               FillStyle       =   0  'Solid
               Height          =   885
               Index           =   1
               Left            =   30
               Top             =   1920
               Width           =   2985
            End
         End
         Begin TabDlg.SSTab SSTab3 
            Height          =   6180
            Left            =   -74910
            TabIndex        =   9
            Top             =   3645
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   10901
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "Armónicos"
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lstArmonicos(1)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Command5(1)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Command4(1)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Combo5(1)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Command3(1)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Command2(1)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Command1(1)"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).ControlCount=   7
            TabCaption(1)   =   "Medición"
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "cmdVelocidad"
            Tab(1).Control(1)=   "cmbEnvolvente(1)"
            Tab(1).Control(2)=   "cmdEnvolvente"
            Tab(1).Control(3)=   "cmdAceleracion"
            Tab(1).Control(4)=   "cmdDesplazamiento"
            Tab(1).Control(5)=   "lstMedicion(1)"
            Tab(1).ControlCount=   6
            TabCaption(2)   =   "Rodamientos"
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "cmdBuscarRodamientos(1)"
            Tab(2).Control(1)=   "txtRodamiento(1)"
            Tab(2).Control(2)=   "cmdRPM(1)"
            Tab(2).Control(3)=   "lstrodamientos(1)"
            Tab(2).Control(4)=   "lblCalFrec(1)"
            Tab(2).Control(5)=   "lblRod(1)"
            Tab(2).Control(6)=   "lblRPM(1)"
            Tab(2).ControlCount=   7
            Begin VB.CommandButton Command1 
               Caption         =   "Max"
               Height          =   315
               Index           =   1
               Left            =   195
               TabIndex        =   23
               Top             =   720
               Width           =   855
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Max 5"
               Height          =   315
               Index           =   1
               Left            =   1125
               TabIndex        =   22
               Top             =   720
               Width           =   855
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Max 10"
               Height          =   315
               Index           =   1
               Left            =   2055
               TabIndex        =   21
               Top             =   720
               Width           =   855
            End
            Begin VB.ComboBox Combo5 
               Height          =   315
               Index           =   1
               Left            =   195
               TabIndex        =   20
               Top             =   1170
               Width           =   1065
            End
            Begin VB.CommandButton Command4 
               Caption         =   "Ir"
               Height          =   315
               Index           =   1
               Left            =   1365
               TabIndex        =   19
               Top             =   1170
               Width           =   495
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Análisis"
               Height          =   315
               Index           =   1
               Left            =   2055
               TabIndex        =   18
               Top             =   1170
               Width           =   855
            End
            Begin VB.CommandButton cmdDesplazamiento 
               Caption         =   "&Desplazamiento"
               Height          =   345
               Left            =   -74820
               TabIndex        =   17
               Top             =   1035
               Width           =   1365
            End
            Begin VB.CommandButton cmdAceleracion 
               Caption         =   "&Aceleración"
               Height          =   345
               Left            =   -74820
               TabIndex        =   16
               Top             =   1485
               Width           =   1365
            End
            Begin VB.CommandButton cmdEnvolvente 
               Caption         =   "&Envolvente"
               Height          =   345
               Left            =   -74820
               TabIndex        =   15
               Top             =   1935
               Width           =   1365
            End
            Begin VB.ComboBox cmbEnvolvente 
               Height          =   315
               Index           =   1
               Left            =   -73320
               TabIndex        =   14
               Top             =   1935
               Width           =   1485
            End
            Begin VB.CommandButton cmdVelocidad 
               Caption         =   "&Velocidad"
               Height          =   345
               Left            =   -74820
               TabIndex        =   13
               Top             =   585
               Width           =   1365
            End
            Begin VB.ComboBox cmdRPM 
               Height          =   315
               Index           =   1
               Left            =   -74220
               TabIndex        =   12
               Top             =   630
               Width           =   1575
            End
            Begin VB.TextBox txtRodamiento 
               Height          =   345
               Index           =   1
               Left            =   -73860
               TabIndex        =   11
               Top             =   1140
               Width           =   1215
            End
            Begin VB.CommandButton cmdBuscarRodamientos 
               Caption         =   "??"
               Height          =   345
               Index           =   1
               Left            =   -72600
               TabIndex        =   10
               Top             =   1140
               Width           =   675
            End
            Begin MSComctlLib.ListView lstArmonicos 
               Height          =   2895
               Index           =   1
               Left            =   135
               TabIndex        =   24
               Top             =   1680
               Width           =   2985
               _ExtentX        =   5265
               _ExtentY        =   5106
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   3
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "No"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Frecuencia"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Amplitud"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView lstMedicion 
               Height          =   2895
               Index           =   1
               Left            =   -74820
               TabIndex        =   25
               Top             =   2445
               Width           =   2955
               _ExtentX        =   5212
               _ExtentY        =   5106
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "PARAMETRO"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Valor"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView lstrodamientos 
               Height          =   2895
               Index           =   1
               Left            =   -74820
               TabIndex        =   26
               Top             =   2100
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   5106
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "DEFECTO"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "FRECUENCIA"
                  Object.Width           =   2540
               EndProperty
            End
            Begin VB.Label lblRPM 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "RPMs"
               Height          =   195
               Index           =   1
               Left            =   -74700
               TabIndex        =   29
               Top             =   690
               Width           =   435
            End
            Begin VB.Label lblRod 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rodamiento"
               Height          =   195
               Index           =   1
               Left            =   -74790
               TabIndex        =   28
               Top             =   1230
               Width           =   855
            End
            Begin VB.Label lblCalFrec 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Calcular Frecuencias"
               Height          =   195
               Index           =   1
               Left            =   -73800
               TabIndex        =   27
               Top             =   1860
               Width           =   1485
            End
         End
         Begin MSComctlLib.TreeView trvListaEmpr 
            Height          =   3945
            Left            =   -74910
            TabIndex        =   62
            Top             =   540
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   6959
            _Version        =   393217
            Style           =   7
            Appearance      =   1
         End
         Begin TabDlg.SSTab XTab1 
            Height          =   6585
            Left            =   -74910
            TabIndex        =   63
            Top             =   4545
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   11615
            _Version        =   393216
            Style           =   1
            Tab             =   1
            TabHeight       =   520
            TabCaption(0)   =   "Lista"
            Tab(0).ControlEnabled=   0   'False
            Tab(0).ControlCount=   0
            TabCaption(1)   =   "Nuevo"
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "frNuevo2"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "frNuevo1"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).ControlCount=   2
            TabCaption(2)   =   "Archivos"
            Tab(2).ControlEnabled=   0   'False
            Tab(2).ControlCount=   0
            Begin VB.Frame frNuevo1 
               BackColor       =   &H00505050&
               Height          =   6195
               Left            =   45
               TabIndex        =   64
               Top             =   450
               Width           =   2955
               Begin VB.CommandButton cmdNuevoSiguiente 
                  BackColor       =   &H00505050&
                  Caption         =   "&Siguiente"
                  Height          =   375
                  Left            =   1680
                  MaskColor       =   &H00505050&
                  TabIndex        =   68
                  Top             =   2100
                  Width           =   1155
               End
               Begin VB.ComboBox cmbEquipo 
                  Height          =   315
                  Left            =   990
                  TabIndex        =   67
                  Text            =   "Combo8"
                  Top             =   1200
                  Width           =   1845
               End
               Begin VB.ComboBox cmbEmpresa 
                  Height          =   315
                  Left            =   990
                  TabIndex        =   66
                  Text            =   "Combo8"
                  Top             =   780
                  Width           =   1845
               End
               Begin VB.ComboBox cmbUbicacion 
                  Height          =   315
                  Left            =   990
                  TabIndex        =   65
                  Text            =   "Combo8"
                  Top             =   360
                  Width           =   1845
               End
               Begin VB.Label lblEquipo 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Equipo"
                  Height          =   195
                  Index           =   1
                  Left            =   390
                  TabIndex        =   71
                  Top             =   1230
                  Width           =   495
               End
               Begin VB.Label lblEmpresa 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Empresa"
                  Height          =   195
                  Index           =   1
                  Left            =   270
                  TabIndex        =   70
                  Top             =   840
                  Width           =   615
               End
               Begin VB.Label lblUbicacion 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ubicación"
                  Height          =   195
                  Index           =   1
                  Left            =   180
                  TabIndex        =   69
                  Top             =   420
                  Width           =   720
               End
            End
            Begin VB.Frame frNuevo2 
               BackColor       =   &H00FFC0C0&
               Height          =   6195
               Left            =   45
               TabIndex        =   72
               Top             =   450
               Width           =   2955
               Begin VB.TextBox txtPAnalisis 
                  Height          =   285
                  Left            =   1740
                  TabIndex        =   83
                  Text            =   "1"
                  Top             =   4320
                  Width           =   1125
               End
               Begin VB.VScrollBar VScroll1 
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   82
                  Top             =   4350
                  Width           =   255
               End
               Begin VB.CommandButton cmdAgregarRod 
                  Caption         =   "&Agregar"
                  Height          =   345
                  Left            =   1740
                  TabIndex        =   81
                  Top             =   5730
                  Width           =   1125
               End
               Begin VB.CommandButton cmdAnterior 
                  Caption         =   "&Anterior"
                  Height          =   345
                  Left            =   300
                  TabIndex        =   80
                  Top             =   5730
                  Width           =   1125
               End
               Begin VB.CommandButton cmdRodamientos 
                  Caption         =   "Rodamientos"
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   79
                  Top             =   5190
                  Width           =   1185
               End
               Begin VB.ComboBox cmdClase 
                  Height          =   315
                  Left            =   480
                  TabIndex        =   78
                  Text            =   "Combo12"
                  Top             =   5190
                  Width           =   675
               End
               Begin VB.ComboBox Combo11 
                  Height          =   315
                  Left            =   1380
                  TabIndex        =   77
                  Text            =   "Combo11"
                  Top             =   4740
                  Width           =   1515
               End
               Begin VB.CommandButton cmdPosterior 
                  Caption         =   ">>"
                  Height          =   345
                  Left            =   1920
                  TabIndex        =   76
                  Top             =   3750
                  Width           =   975
               End
               Begin VB.CommandButton cmdPrevio 
                  Caption         =   "<<"
                  Height          =   345
                  Left            =   780
                  TabIndex        =   75
                  Top             =   3750
                  Width           =   975
               End
               Begin VB.PictureBox picInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Left            =   510
                  ScaleHeight     =   225
                  ScaleWidth      =   150
                  TabIndex        =   74
                  Top             =   3810
                  Width           =   150
               End
               Begin VB.PictureBox picNivSev 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Left            =   1290
                  ScaleHeight     =   225
                  ScaleWidth      =   150
                  TabIndex        =   73
                  Top             =   5220
                  Width           =   150
               End
               Begin VB.Image imgPM 
                  Height          =   3345
                  Left            =   60
                  Stretch         =   -1  'True
                  Top             =   180
                  Width           =   2835
               End
               Begin VB.Label lblClase 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Clase"
                  Height          =   195
                  Index           =   1
                  Left            =   30
                  TabIndex        =   86
                  Top             =   5250
                  Width           =   390
               End
               Begin VB.Label lblAcopl 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Acoplamiento"
                  Height          =   195
                  Index           =   1
                  Left            =   270
                  TabIndex        =   85
                  Top             =   4770
                  Width           =   960
               End
               Begin VB.Label lblPAnalisis 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Puntos de Análisis"
                  Height          =   195
                  Index           =   1
                  Left            =   210
                  TabIndex        =   84
                  Top             =   4350
                  Width           =   1290
               End
            End
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3555
      Top             =   2565
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":020A
            Key             =   "imprimir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":0C1C
            Key             =   "ayuda"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":811E
            Key             =   "nuevo"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDevice 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DISPOSITIVO DE ADQUISICIÓN DE DATOS NO CONECTADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6690
      TabIndex        =   6
      Top             =   1800
      Width           =   6390
   End
   Begin VB.Image imgSplitter 
      Height          =   2100
      Left            =   3270
      MousePointer    =   9  'Size W E
      Top             =   450
      Width           =   45
   End
   Begin VB.Menu mnuFic 
      Caption         =   "&Ficheros"
      Begin VB.Menu mnuFormenPanelIzq 
         Caption         =   "Mostrar el Form2 en el panel &izquierdo"
      End
      Begin VB.Menu mnuFormenPanelDer 
         Caption         =   "Mostrar el Form2 en el panel &derecho"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNotepad 
         Caption         =   "Mostrar el &bloc de notas"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuVentanas 
      Caption         =   "&Ventanas"
      Begin VB.Menu mnuV_Cerrar 
         Caption         =   "&Cerrar"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mnuAcercade 
         Caption         =   "Acerca de..."
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Tooltips As New Collection
Dim flagFig As Integer
'------------------------------------------------------------------------------
Private cn As ADODB.connection
Private objIzq As PictureBox
Private objDer As PictureBox
Private objSup As PictureBox
Private objInf As PictureBox
Private moviendo As Boolean
Private Const splitLimit As Long = 15&
Private hndNotepad As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'
'------------------------------------------------------------------------------
' APIS para incluir las ventanas en un PictureBox
'------------------------------------------------------------------------------
' Para hacer ventanas hijas
'Dim prevParent As Long
Private Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
' Para mostrar una ventana según el handle (hwnd)
' ShowWindow() Commands
Private Enum eShowWindow
    HIDE_eSW = 0&
    SHOWNORMAL_eSW = 1&
    NORMAL_eSW = 1&
    SHOWMINIMIZED_eSW = 2&
    SHOWMAXIMIZED_eSW = 3&
    MAXIMIZE_eSW = 3&
    SHOWNOACTIVATE_eSW = 4&
    SHOW_eSW = 5&
    MINIMIZE_eSW = 6&
    SHOWMINNOACTIVE_eSW = 7&
    SHOWNA_eSW = 8&
    RESTORE_eSW = 9&
    SHOWDEFAULT_eSW = 10&
    MAX_eSW = 10&
End Enum

Private Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, ByVal nCmdShow As eShowWindow) As Long
' Para posicionar una ventana según su hWnd
Private Declare Function MoveWindow Lib "User32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
' Para saber si una ventana es hija de otra
Private Declare Function IsChild Lib "User32" (ByVal hWndParent As Long, ByVal hWnd As Long) As Long
'
Private Type PointAPI
    X As Long
    Y As Long
End Type
Private Type RECTAPI
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    ShowCmd As Long
    ptMinPosition As PointAPI
    ptMaxPosition As PointAPI
    rcNormalPosition As RECTAPI
End Type
Private Declare Function GetWindowPlacement Lib "User32" (ByVal hWnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Long

'------------------------------------------------------------------------------
' Procedimientos NO de evento
'------------------------------------------------------------------------------
' Mostrar el formulario indicado, dentro de picDock
Private Sub dockForm(ByVal formhWnd As Long, ByVal picDock As PictureBox, Optional ByVal ajustar As Boolean = True)
    ' Hacer el formulario indicado, un hijo del picDock
    ' Si Ajustar es True, se ajustará al tamaño del contenedor,
    ' si Ajustar es False, se quedará con el tamaño actual.
    Call SetParent(formhWnd, picDock.hWnd)
    posDockForm formhWnd, picDock, ajustar
    Call ShowWindow(formhWnd, NORMAL_eSW)
End Sub

' Posicionar el formulario indicado dentro de picDock
Private Sub posDockForm(ByVal formhWnd As Long, ByVal picDock As PictureBox, Optional ByVal ajustar As Boolean = True)
    ' Posicionar el formulario indicado en las coordenadas del picDock
    ' Si Ajustar es True, se ajustará al tamaño del contenedor,
    ' si Ajustar es False, se quedará con el tamaño actual.
    Dim nWidth As Long, nHeight As Long
    Dim wndPl As WINDOWPLACEMENT
    '
   'On Error GoTo Err_Proc

    If ajustar Then
        nWidth = picDock.ScaleWidth \ Screen.TwipsPerPixelX
        nHeight = picDock.ScaleHeight \ Screen.TwipsPerPixelY
    Else
        ' el tamaño del formulario que se va a posicionar
        Call GetWindowPlacement(formhWnd, wndPl)
        With wndPl.rcNormalPosition
            nWidth = .Right - .Left
            nHeight = .Bottom - .Top
        End With
    End If
    Call MoveWindow(formhWnd, 0, 0, nWidth, nHeight, True)

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPrincipal", "posDockForm"
   Err.Clear
   Resume Exit_Proc

End Sub

' Este procedimiento se usará para ajustar los tamaños de los paneles
Private Sub sizeControls(ByVal X As Long)
    Dim tMinWidth As Long
    '
    On Error Resume Next
    ' el ancho mínimo que tendrá cada panel
    tMinWidth = Screen.TwipsPerPixelY * 90
    ' asignar el ancho
    If X < tMinWidth Then X = tMinWidth
    If X > (Me.Width - tMinWidth) Then X = Me.Width - tMinWidth
    objIzq.Width = X
    imgSplitter.Left = X
    objDer.Left = X + 90
    objDer.Width = Me.ScaleWidth - (objIzq.Width + imgSplitter.Width)  '140)
    ' asignar la parte superior
    ' aquí se puede usar otro control para saber dónde situar la parte superior
    objIzq.Top = objSup.Top + objSup.Height
    objDer.Top = objIzq.Top
    ' asignar la altura
    ' aquí se puede usar otro control para saber dónde situar la parte inferior
    If objInf.Visible Then
        objIzq.Height = Me.ScaleHeight - (objSup.Top + objSup.Height + objInf.Height)
    Else
        objIzq.Height = Me.ScaleHeight - (objSup.Top + objSup.Height)
    End If
    '
    objDer.Height = objIzq.Height
    imgSplitter.Top = objIzq.Top
    imgSplitter.Height = objIzq.Height
End Sub

Private Sub Form_Activate()
Me.WindowState = vbMaximized
End Sub

'------------------------------------------------------------------------------
' Procedimientos de evento
'------------------------------------------------------------------------------
Private Sub Form_Load()
Dim Tooltip   As cToolTip
    ' asignar los controles que realmente se usarán
    ' En caso de que no sean Pictures, cambiarlos en la declaración
   'On Error GoTo Err_Proc
Leer_Escribir_Timer
    Set objIzq = picL
    Set objDer = picR
    Set objSup = picT
    'Set objInf = stBar   'picD
    '
    ' asignar el "acoplamiento"
    objIzq.Align = vbAlignLeft
    objDer.Align = vbAlignRight
    '
    picSplitter.Visible = False
    '
    Me.WindowState = vbMaximized
    
    ' mostrar el segundo formulario e incluirlo en el picDock
    '
    ' Asignar el Tag del formulario para que se ajuste al tamaño del objIzq
    ''frmTabMain.Tag = "objIzq"
    '
    ''dockForm frmTabMain.hWnd, objIzq, True
'descripción indvidual
Set Tooltip = New cToolTip
Tooltip.Create picInfo, "Por norma, se establece el número del punto de medición contando a partir del lado libre del motor|en dirección del lado del impulsor.|    1.- Motor lado libre. |    2.- Motor lado acople. |    3.- Impulsor lado acople.|    4.- Impulsor lado libre.", _
                             TTBalloonIfActive, False, TTIconInfo, "Puntos de Medición", vbBlack, vbWhite, 100, 20000
Tooltips.Add Tooltip, picInfo.Name 'no te olvides de mantenerlo
flagFig = 1

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPrincipal", "Form_Load"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next

Dim tForm As Form
Dim s As String
Dim Msg   ' Declara la variable.
' Establece el texto del mensaje.
Msg = "¿Realmente desea salir de la aplicación?"
' Si el usuario hace clic en el botón No, se detiene QueryUnload.
If MsgBox(Msg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then
    Cancel = True
Else
    s = Me.Name
    For Each tForm In Forms
        If tForm.Name <> s Then
            Unload tForm
        End If
    Next
    Set Tooltips = Nothing
    End
End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    'lblDevice.Left = Me.ScaleWidth - lblDevice.Width - 120
    'stBarFrame1.Left = lblDevice.Left - 50
    ' Los tamaños mínimos del formulario
    If Me.Width < 6000 Then Me.Width = 6000
    If Me.Height < 3000 Then Me.Height = 3000
    ' ajustar los controles al nuevo tamaño
    sizeControls imgSplitter.Left
    'picL
    tabMain.Height = picL.Height - 5
    tabMain.Width = picL.Width - 20
    frPlano1.Width = tabMain.Width - 50
    frPlano2.Width = tabMain.Width - 50
    lblPlano(1).Left = frPlano1.Left + 5: lblPlano(2).Left = frPlano2.Left + 5
    Err = 0
End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'On Error GoTo Err_Proc

    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 3, .Height - 20
    End With
    picSplitter.Visible = True
    moviendo = True

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPrincipal", "imgSplitter_MouseDown"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    '
   'On Error GoTo Err_Proc

    If moviendo Then
        sglPos = X + imgSplitter.Left
        If sglPos < splitLimit Then
            picSplitter.Left = splitLimit
        ElseIf sglPos > Me.Width - splitLimit Then
            picSplitter.Left = Me.Width - splitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPrincipal", "imgSplitter_MouseMove"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sizeControls picSplitter.Left
    picSplitter.Visible = False
    moviendo = False
End Sub

Private Sub Label1_Click()
Leer_Escribir_Timer
End Sub

Private Sub mnuAcercade_Click()
  'Carga y visualiza el formulario Splash
   'On Error GoTo Err_Proc

  Load frmAcercade   'frmSplash
  ' Carga en memoria el formulario principal pero no lo muestra
  Load frmPrincipal
  ' ..Hasta que no se cumpla el tiempo se visualiza el Splash
  Do
    DoEvents
  Loop Until frmAcercade.Listo
  ' descarga el Splash con una animación
  Call Animar(frmAcercade, 500, AW_BLEND Or AW_HIDE)
  Unload frmAcercade
  Set frmAcercade = Nothing

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPrincipal", "mnuAcercade_Click"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub mnuFormenPanelDer_Click()
    ' mostrar el segundo formulario e incluirlo en el picDock
    ' Asignar el Tag del formulario para que se ajuste al tamaño del objDer
    frmBase.Tag = "objDer"
    '
    dockForm frmBase.hWnd, objDer, True
    'frmPloteaSeñal
End Sub

Private Sub mnuFormenPanelIzq_Click()
    ' mostrar el segundo formulario e incluirlo en el picDock
    ' Asignar el Tag del formulario para que se ajuste al tamaño del objIzq
    'frmTabMain.Tag = "objIzq"
    '
    'dockForm frmTabMain.hWnd, objIzq, True
End Sub

Private Sub mnuNotepad_Click()
    ' mostrar el block de notas en el panel izquierdo
    '
    Call Shell("notepad.exe", vbNormalFocus)
    '
    Dim s As String
    '**************************************************************************
    ' NOTA:                                                         (26/May/04)
    '   Si estás usando un Windows con la versión en español,
    '   tendrás que cambiar "Untitled - Notepad" por el correspondiente nombre,
    '   creo que es "Sin Título - Bloc de notas"
    '**************************************************************************
    hndNotepad = FindWindow(s, "Sin Título - Bloc de notas")
    '
    dockForm hndNotepad, objDer, True
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub

Private Sub mnuV_Cerrar_Click()
    Unload frmBase
End Sub

Private Sub picL_DragDrop(Source As control, X As Single, Y As Single)
    If Source = imgSplitter Then
        sizeControls X
    End If
End Sub

Private Sub picL_Resize()
    ' Posicionar los formularios "hijos" de este control
    Dim tForm As Form
    '
   'On Error GoTo Err_Proc

    For Each tForm In Forms
        ' El tag del formulario incluido en el picture tendrá el Tag asignado
        ' con el valor "objIzq"
        If CStr(tForm.Tag) = "objIzq" Then
            posDockForm tForm.hWnd, objIzq
        End If
    Next

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPrincipal", "picL_Resize"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub picR_Resize()
    ' Posicionar los formularios "hijos" de este control
    Dim tForm As Form
    '
    For Each tForm In Forms
        ' El tag del formulario incluido en el picture tendrá el Tag asignado
        ' con el valor "objDer"
        If CStr(tForm.Tag) = "objDer" Then
            posDockForm tForm.hWnd, objDer
        End If
    Next
    '
    On Error Resume Next
    If hndNotepad > 0 Then
        posDockForm hndNotepad, objDer
    End If
End Sub

Sub ChequeaOrtografia()
Dim objWord As Object
    Dim objDoc  As Object
    Dim strResult As String
    
    'Crear una nueva instanacia de Word
   'On Error GoTo Err_Proc

    Set objWord = CreateObject("word.Application")

    Select Case objWord.Version
        'Office 2000
        Case "9.0"
            Set objDoc = objWord.Documents.Add(, , 1, True)
        'Office XP
        Case "10.0"
            Set objDoc = objWord.Documents.Add(, , 1, True)
        'Office 97
        Case Else ' Office 97
            Set objDoc = objWord.Documents.Add
    End Select

    objDoc.Content = Text1.Text
    objDoc.CheckSpelling

    strResult = Left(objDoc.Content, Len(objDoc.Content) - 1)

    If Text1.Text = strResult Then
         'No hubo errores de ortografía, por lo que le da al usuario y
         'Señal visual de que algo pasó
        MsgBox "The spelling check is complete.", vbInformation + vbOKOnly
    End If
    
    'Clean up
    objDoc.Close False
    Set objDoc = Nothing
    objWord.Application.Quit True
    Set objWord = Nothing

    'Vuelva a colocar el texto seleccionado con el texto corregido. Es así de importante
    'Sea este hecho después de que el" Clean Up "Porque de lo contrario hay problemas
    'Con la pantalla no volver a pintar
    Text1.Text = strResult

    Exit Sub

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPrincipal", "ChequeaOrtografia"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub ColorListviewRow(lv As ListView, RowNbr As Long, RowColor As OLE_COLOR)
'***************************************************************************
'Proposito: El color y la fila del ListView
'Entradas : lv - El listview
'           RowNbr - El indice de la fila a colorear
'           RowColor - El color para colorear
'Salidas  : Ninguna
'***************************************************************************
    
    Dim itmX As ListItem
    Dim lvSI As ListSubItem
    Dim intIndex As Integer
    
    On Error GoTo ErrorRoutine
    
    Set itmX = lv.ListItems(RowNbr)
    itmX.ForeColor = RowColor
    For intIndex = 1 To lv.ColumnHeaders.Count - 1
        Set lvSI = itmX.ListSubItems(intIndex)
        lvSI.ForeColor = RowColor
    Next

    Set itmX = Nothing
    Set lvSI = Nothing
    
    Exit Sub

ErrorRoutine:

    MsgBox Err.Description

End Sub


'Rutinas que inician la conexion USB basada en un timer que testea permanentemente
'los conectores USB de la computadora.
Private Sub Leer_Escribir_Timer()
'On Error Resume Next
Dim output, I As Byte
Dim conectado As Boolean
'Testea si el dispositivo esta conectado
If MyDeviceDetected Then
    'stBar.Panels.Add , "clave1", "DISPOSITIVO DE ADQUISICIÓN DE DATOS CONECTADO", sbrText
    stBar.Panels(1).Style = sbrText
    stBar.Panels(1).Text = "DISPOSITIVO DE ADQUISICIÓN DE DATOS CONECTADO"
    stBar.Panels(1).Picture = LoadPicture("C:\VibraMec\Iconos\wxp\10.ico")
    stBar.Panels(1).Style = sbrText
    stBar.Panels(1).Width = 6000
    stBar.Refresh
Else
    'stBar.Panels.Add , "clave1", "DISPOSITIVO DE ADQUISICIÓN DE DATOS NO CONECTADO", sbrText
    stBar.Panels(1).Style = sbrText
    stBar.Panels(1).Text = "DISPOSITIVO DE ADQUISICIÓN DE DATOS NO CONECTADO"
    stBar.Panels(1).Picture = LoadPicture("C:\Vibramek\VibrameK\wxp\11.ico")
    stBar.Panels(1).Style = sbrText
    stBar.Panels(1).Width = 6000
    stBar.Refresh
    conectado = False
    FindTheHid
End If

'Lectura/escritura del dispositivo HID
'Call ReadAndWriteToDevice
'Indica el nivel de entrada en una barra de progreso
'ProgressBar1(0).Value = ReadBuffer(1) + ReadBuffer(2) * 256
'ProgressBar1(1).Value = ReadBuffer(3) + ReadBuffer(4) * 256
'ProgressBar1(2).Value = ReadBuffer(5) + ReadBuffer(6) * 256
'tomamos el segundo MSB del ADC en el buffer (8) e ignora las 4 entradas digitales del PIC
'ProgressBar1(3).Value = ReadBuffer(7) + (ReadBuffer(8) And 3) * 256
'Los bits del ADC seran removidos solo si los 4 digitos de las entradas del PIC estan en cero.
'output = ReadBuffer(8) / 16
'Indica si las entradas digitales en el PIC estan activas y cambia el color del shape indicando el estado.
'For i = 1 To 4
    'If output Mod (2 ^ i) > ((2 ^ i) / 2 - 1) Then
        'Shape1(i - 1).FillColor = &HFF00&
    'Else
        'Shape1(i - 1).FillColor = &HFF&
    'End If
'Next
'Envía un byte a la salida digital RB0 del PIC.
'OutputReportData(1) = Check1(0).Value + Check1(1).Value * 2 + Check1(2).Value * 4 + Check1(3).Value * 8

End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cmdAnterior_Click()
    frNuevo2.Visible = False
    frNuevo1.Visible = True
End Sub

Private Sub cmdNuevoSiguiente_Click()
    frNuevo1.Visible = False
    frNuevo2.Visible = True
End Sub

Private Sub cmdAgregarRod_Click()
    frNuevo2.Visible = False
    frNuevo1.Visible = True
End Sub

Private Sub cmdPosterior_Click()
If flagFig >= 8 Then
    flagFig = 8
    Set imgPM.Picture = LoadPicture(App.Path & "\Imagenes\Fig_" & flagFig & ".jpg", vbLPLarge, vbLPColor)
Else
    flagFig = flagFig + 1
    Set imgPM.Picture = LoadPicture(App.Path & "\Imagenes\Fig_" & flagFig & ".jpg", vbLPLarge, vbLPColor)
End If
End Sub

Private Sub cmdPrevio_Click()
If flagFig <= 1 Then
    flagFig = 1
    Set imgPM.Picture = LoadPicture(App.Path & "\Imagenes\Fig_" & flagFig & ".jpg", vbLPLarge, vbLPColor)
Else
    flagFig = flagFig - 1
    Set imgPM.Picture = LoadPicture(App.Path & "\Imagenes\Fig_" & flagFig & ".jpg", vbLPLarge, vbLPColor)
End If
'imgLstPM
End Sub

Private Sub lblPlano_Click(Index As Integer)
If lblPlano(Index).Index = 1 Then
    lblPlano(1).BackColor = vbRed
    lblPlano(1).ForeColor = &HC0FFC0
    lblPlano(2).BackColor = &HFF0000
    lblPlano(2).ForeColor = &H8000&
Else
    lblPlano(2).BackColor = vbRed
    lblPlano(2).ForeColor = &HC0FFC0
    lblPlano(1).BackColor = &HFF0000
    lblPlano(1).ForeColor = &H8000&
End If
End Sub

Private Sub picNivSev_Click()
    frmNiv_Sev.Show vbModal
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------
'Carga los registros al hacer Click en una tabla
Private Sub Cargar_Registro(ByVal Nodo_Tabla As Node)
On Local Error GoTo ErrSub
Dim rst As ADODB.Recordset
Dim I As Integer
Dim RStr As String
Dim Nodo As Node
    ' Si ya había cargado sale
    If Nodo_Tabla.Children > 0 Then Exit Sub
    ' Cantidad de registros
    Set rst = cn.Execute("SELECT Count(*) as Total FROM " & Nodo_Tabla.Text, , adCmdText)
    'ProgressBar1.Max = rst("Total")
    'ProgressBar1.value = 0
    ' Llena el recordset con los registros
    Set rst = cn.Execute("SELECT * FROM " & Nodo_Tabla.Text, , adCmdText)
    Me.MousePointer = vbHourglass
    ' Recorre el recordset para añadir en los registros al Nodo del TreeView
    Do Until rst.EOF
        RStr = ""
        'Recorre los campos
        For I = 0 To rst.Fields.Count - 1
            RStr = RStr & ", " & rst.Fields.item(I)
        Next I
        
        RStr = Mid$(RStr, 2)

        Set Nodo = trvListaEmpr.Nodes.Add(Nodo_Tabla, tvwChild, , RStr)
        ' Tag del nodo
        Nodo.Tag = "Registro"
        ' siguiente registro
        rst.MoveNext
        'Progreso
        'ProgressBar1.value = ProgressBar1.value + 1
    Loop
    Me.MousePointer = vbDefault
    rst.Close 'Cierra el recordset
Exit Sub
'Error
ErrSub:
MsgBox Err.Description, vbCritical
End Sub

' Carga las tablas en el TreeView
Private Sub Cargar_Tabla()
On Local Error GoTo ErrSub
Dim rst As ADODB.Recordset
Dim Nodo_Tabla As Node
    Set rst = cn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "Table"))
    Do While Not rst.EOF
        Set Nodo_Tabla = trvListaEmpr.Nodes.Add(, , , rst!TABLE_NAME)
        Nodo_Tabla.Tag = "Tabla"
        rst.MoveNext
    Loop
    rst.Close
Exit Sub
ErrSub:
MsgBox Err.Description, vbCritical
End Sub

Private Sub Abrir_Base_Dato(path_bd As String)
    ' Nuevo objeto ADODB.Connection
    Set cn = New ADODB.connection
    'Cadena de conexión
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & path_bd & ";" & "Persist Security Info=False"
    'Abre la base de datos
    cn.Open
    ' Elimina los elementos del TreeView
    trvListaEmpr.Nodes.Clear
    ' Carga las Tablas en los nodos del TreeView
    Cargar_Tabla
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "nuevoBal1"
        Call mnuFormenPanelDer_Click
End Select
End Sub

Private Sub trvListaEmpr_NodeClick(ByVal Node As MSComctlLib.Node)
    Select Case Node.Tag
        ' Si se hizo click en una Tabla, carga los registros en el nodo
        Case "Tabla": Call Cargar_Registro(Node)
        ' Si se hizo click en un registro, muestra el dato
        Case "Registro": MsgBox "Registro: " & Node.Text, vbInformation
    End Select
End Sub

'Botón para abrir la bd
'Private Sub Command1_Click()
'With CommonDialog1
'     .Filter = "Archivos MDB|*.mdb"
'     .DialogTitle = " Seleccionar base de datos "
'     .ShowOpen
'     If .FileName = "" Then Exit Sub
'     ' Le pasa el Path de la base de datos
'     Call Abrir_Base_Dato(.FileName)
'     Me.Caption = .FileName
'End With
'End Sub

'Private Sub Form_Load()
'Me.ScaleMode = vbTwips
'ProgressBar1.Height = 375
'With Command1
'    .Height = 375
'    .Width = 2000
'    Command1.Caption = " ->> Abrir base "
'End With
'End Sub

'Private Sub Form_Resize()
'    trvListaEmpr.Move 0, 0, ScaleWidth, ScaleHeight - 375
'    ProgressBar1.Move 0, ScaleHeight - 375, ScaleWidth - Command1.Width
'    Command1.Move ScaleWidth - Command1.Width, ScaleHeight - 375
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    On Local Error Resume Next
'    cn.Close
'End Sub





