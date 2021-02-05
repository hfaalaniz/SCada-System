VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBase 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Osciloscopio"
   ClientHeight    =   12585
   ClientLeft      =   4770
   ClientTop       =   900
   ClientWidth     =   22005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12585
   ScaleWidth      =   22005
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPDer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawWidth       =   2
      ForeColor       =   &H00808080&
      Height          =   4500
      Left            =   4500
      Picture         =   "frmBase.frx":0000
      ScaleHeight     =   4470
      ScaleWidth      =   4470
      TabIndex        =   235
      Top             =   495
      Width           =   4500
   End
   Begin VB.PictureBox picPIzq 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawWidth       =   2
      ForeColor       =   &H00808080&
      Height          =   4500
      Left            =   0
      Picture         =   "frmBase.frx":40E9A
      ScaleHeight     =   4470
      ScaleWidth      =   4470
      TabIndex        =   234
      Top             =   495
      Width           =   4500
   End
   Begin VB.PictureBox picPolar1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   14040
      ScaleHeight     =   4470
      ScaleWidth      =   4470
      TabIndex        =   229
      Top             =   6435
      Width           =   4500
   End
   Begin VB.HScrollBar VH1 
      Height          =   240
      Left            =   11835
      Max             =   3000
      TabIndex        =   226
      Top             =   5580
      Width           =   4515
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   45
      TabIndex        =   33
      Top             =   5895
      Width           =   19140
      _ExtentX        =   33761
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Inicio"
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grpTendencias"
      Tab(0).Control(1)=   "HScroll1"
      Tab(0).Control(2)=   "Label1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Balanceo"
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Scope(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Scope(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Calculadora"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab2"
      Tab(2).ControlCount=   1
      Begin VibraMec.aGraph grpTendencias 
         Height          =   3840
         Left            =   -74865
         TabIndex        =   233
         Top             =   1260
         Width           =   13785
         _ExtentX        =   24315
         _ExtentY        =   6773
         GraphLineColor  =   0
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   285
         Left            =   -74820
         Max             =   500
         TabIndex        =   227
         Top             =   810
         Value           =   250
         Width           =   5190
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4785
         Left            =   -74910
         TabIndex        =   36
         Top             =   450
         Width           =   18915
         _ExtentX        =   33364
         _ExtentY        =   8440
         _Version        =   393216
         Style           =   1
         Tabs            =   9
         Tab             =   8
         TabsPerRow      =   9
         TabHeight       =   520
         TabCaption(0)   =   "Balanceo en 1 Plano"
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame1"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Balanceo en 2 Planos"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame6"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Balanceo sin Fase"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Image9"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "División del Peso"
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame7"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Unificar Masas"
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame11"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Mechas"
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame14"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Placas"
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Frame16"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "Peso de Prueba"
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "Frame18"
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "Radio de los Pesos"
         Tab(8).ControlEnabled=   -1  'True
         Tab(8).Control(0)=   "Frame20"
         Tab(8).Control(0).Enabled=   0   'False
         Tab(8).ControlCount=   1
         Begin VB.Frame Frame20 
            Caption         =   "Radio de Peso"
            Height          =   2595
            Left            =   450
            TabIndex        =   213
            Top             =   1035
            Width           =   5325
            Begin VB.VScrollBar VScroll40 
               Height          =   285
               Left            =   4920
               TabIndex        =   222
               Top             =   1320
               Width           =   195
            End
            Begin VB.VScrollBar VScroll41 
               Height          =   285
               Left            =   4920
               TabIndex        =   221
               Top             =   390
               Width           =   195
            End
            Begin VB.VScrollBar VScroll42 
               Height          =   285
               Left            =   4920
               TabIndex        =   220
               Top             =   870
               Width           =   195
            End
            Begin VB.CommandButton Command10 
               Caption         =   "Calcular"
               Height          =   315
               Left            =   1740
               TabIndex        =   219
               Top             =   2040
               Width           =   1365
            End
            Begin VB.Frame Frame21 
               Height          =   615
               Left            =   3690
               TabIndex        =   217
               Top             =   1800
               Width           =   1425
               Begin VB.Label Label51 
                  AutoSize        =   -1  'True
                  Caption         =   "Resultado"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   150
                  TabIndex        =   218
                  Top             =   240
                  Width           =   930
               End
            End
            Begin VB.TextBox Text40 
               Height          =   315
               Left            =   3660
               TabIndex        =   216
               Top             =   390
               Width           =   1485
            End
            Begin VB.TextBox Text41 
               Height          =   315
               Left            =   3660
               TabIndex        =   215
               Top             =   870
               Width           =   1485
            End
            Begin VB.TextBox Text42 
               Height          =   315
               Left            =   3660
               TabIndex        =   214
               Top             =   1320
               Width           =   1485
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               Caption         =   "Peso (grs)"
               Height          =   195
               Left            =   2880
               TabIndex        =   225
               Top             =   1380
               Width           =   705
            End
            Begin VB.Label Label53 
               AutoSize        =   -1  'True
               Caption         =   "Distancia del centro"
               Height          =   195
               Left            =   2190
               TabIndex        =   224
               Top             =   450
               Width           =   1410
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               Caption         =   "Nueva distancia del centro"
               Height          =   195
               Left            =   1680
               TabIndex        =   223
               Top             =   930
               Width           =   1905
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "Peso de Prueba"
            Height          =   4305
            Left            =   -74595
            TabIndex        =   190
            Top             =   315
            Width           =   5325
            Begin VB.TextBox Text39 
               Height          =   315
               Left            =   3660
               TabIndex        =   207
               Top             =   1800
               Width           =   1485
            End
            Begin VB.TextBox Text38 
               Height          =   315
               Left            =   3660
               TabIndex        =   206
               Top             =   1320
               Width           =   1485
            End
            Begin VB.TextBox Text37 
               Height          =   315
               Left            =   3660
               TabIndex        =   205
               Top             =   870
               Width           =   1485
            End
            Begin VB.TextBox Text36 
               Height          =   315
               Left            =   3660
               TabIndex        =   204
               Top             =   390
               Width           =   1485
            End
            Begin VB.VScrollBar VScroll36 
               Height          =   285
               Left            =   4920
               TabIndex        =   203
               Top             =   1320
               Width           =   195
            End
            Begin VB.VScrollBar VScroll37 
               Height          =   285
               Left            =   4920
               TabIndex        =   202
               Top             =   390
               Width           =   195
            End
            Begin VB.VScrollBar VScroll38 
               Height          =   285
               Left            =   4920
               TabIndex        =   201
               Top             =   870
               Width           =   195
            End
            Begin VB.CommandButton Command9 
               Caption         =   "Calcular"
               Height          =   315
               Left            =   180
               TabIndex        =   200
               Top             =   3840
               Width           =   1365
            End
            Begin VB.Frame Frame19 
               Height          =   615
               Left            =   1860
               TabIndex        =   197
               Top             =   3570
               Width           =   3255
               Begin VB.Label Label43 
                  AutoSize        =   -1  'True
                  Caption         =   "Resultado"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   2010
                  TabIndex        =   199
                  Top             =   240
                  Width           =   930
               End
               Begin VB.Label Label44 
                  AutoSize        =   -1  'True
                  Caption         =   "Resultado"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   150
                  TabIndex        =   198
                  Top             =   240
                  Width           =   930
               End
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               Left            =   3450
               TabIndex        =   196
               Text            =   "Combo1"
               Top             =   2880
               Width           =   1695
            End
            Begin VB.OptionButton Option12 
               Caption         =   "Milímetros"
               Height          =   285
               Left            =   2280
               TabIndex        =   195
               Top             =   3300
               Width           =   1125
            End
            Begin VB.OptionButton Option13 
               Caption         =   "Pulgadas"
               Height          =   255
               Left            =   4110
               TabIndex        =   194
               Top             =   3300
               Width           =   1005
            End
            Begin VB.VScrollBar VScroll39 
               Height          =   285
               Left            =   4920
               TabIndex        =   193
               Top             =   1800
               Width           =   195
            End
            Begin VB.CheckBox Check3 
               Caption         =   "2 Planos"
               Height          =   285
               Left            =   3660
               TabIndex        =   192
               Top             =   2220
               Width           =   1005
            End
            Begin VB.CheckBox Check4 
               Caption         =   "Base Flotante"
               Height          =   285
               Left            =   3660
               TabIndex        =   191
               Top             =   2520
               Width           =   1305
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               Caption         =   "Ancho de placa (mm)"
               Height          =   195
               Left            =   2130
               TabIndex        =   212
               Top             =   1380
               Width           =   1500
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               Caption         =   "Peso de corrección (grs)"
               Height          =   195
               Left            =   1890
               TabIndex        =   211
               Top             =   450
               Width           =   1725
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               Caption         =   "Espesor de la placa (mm)"
               Height          =   195
               Left            =   1860
               TabIndex        =   210
               Top             =   930
               Width           =   1770
            End
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               Caption         =   "Material"
               Height          =   195
               Left            =   2730
               TabIndex        =   209
               Top             =   2970
               Width           =   555
            End
            Begin VB.Label Label49 
               AutoSize        =   -1  'True
               Caption         =   "Ancho de placa (mm)"
               Height          =   195
               Left            =   2130
               TabIndex        =   208
               Top             =   1860
               Width           =   1500
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "Placas"
            Height          =   3855
            Left            =   -74325
            TabIndex        =   172
            Top             =   630
            Width           =   5325
            Begin VB.TextBox Text35 
               Height          =   315
               Left            =   3660
               TabIndex        =   185
               Top             =   1320
               Width           =   1485
            End
            Begin VB.TextBox Text34 
               Height          =   315
               Left            =   3660
               TabIndex        =   184
               Top             =   870
               Width           =   1485
            End
            Begin VB.TextBox Text33 
               Height          =   315
               Left            =   3660
               TabIndex        =   183
               Top             =   390
               Width           =   1485
            End
            Begin VB.OptionButton Option10 
               Caption         =   "Pulgadas"
               Height          =   255
               Left            =   4110
               TabIndex        =   182
               Top             =   2580
               Width           =   1005
            End
            Begin VB.OptionButton Option11 
               Caption         =   "Milímetros"
               Height          =   285
               Left            =   2280
               TabIndex        =   181
               Top             =   2580
               Width           =   1125
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               Left            =   3450
               TabIndex        =   180
               Text            =   "Combo1"
               Top             =   2040
               Width           =   1695
            End
            Begin VB.Frame Frame17 
               Height          =   615
               Left            =   1860
               TabIndex        =   177
               Top             =   3000
               Width           =   3255
               Begin VB.Label Label37 
                  AutoSize        =   -1  'True
                  Caption         =   "Resultado"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   150
                  TabIndex        =   179
                  Top             =   240
                  Width           =   930
               End
               Begin VB.Label Label38 
                  AutoSize        =   -1  'True
                  Caption         =   "Resultado"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   2010
                  TabIndex        =   178
                  Top             =   240
                  Width           =   930
               End
            End
            Begin VB.CommandButton Command8 
               Caption         =   "Calcular"
               Height          =   315
               Left            =   180
               TabIndex        =   176
               Top             =   3270
               Width           =   1365
            End
            Begin VB.VScrollBar VScroll33 
               Height          =   285
               Left            =   4920
               TabIndex        =   175
               Top             =   870
               Width           =   195
            End
            Begin VB.VScrollBar VScroll34 
               Height          =   285
               Left            =   4920
               TabIndex        =   174
               Top             =   390
               Width           =   195
            End
            Begin VB.VScrollBar VScroll35 
               Height          =   285
               Left            =   4920
               TabIndex        =   173
               Top             =   1320
               Width           =   195
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               Caption         =   "Material"
               Height          =   195
               Left            =   2730
               TabIndex        =   189
               Top             =   2130
               Width           =   555
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               Caption         =   "Espesor de la placa (mm)"
               Height          =   195
               Left            =   1860
               TabIndex        =   188
               Top             =   930
               Width           =   1770
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               Caption         =   "Peso de corrección (grs)"
               Height          =   195
               Left            =   1890
               TabIndex        =   187
               Top             =   450
               Width           =   1725
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               Caption         =   "Ancho de placa (mm)"
               Height          =   195
               Left            =   2130
               TabIndex        =   186
               Top             =   1380
               Width           =   1500
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Mechas"
            Height          =   3135
            Left            =   -74460
            TabIndex        =   157
            Top             =   675
            Width           =   5325
            Begin VB.TextBox Text32 
               Height          =   315
               Left            =   3660
               TabIndex        =   168
               Top             =   870
               Width           =   1485
            End
            Begin VB.TextBox Text31 
               Height          =   315
               Left            =   3660
               TabIndex        =   167
               Top             =   390
               Width           =   1485
            End
            Begin VB.VScrollBar VScroll31 
               Height          =   285
               Left            =   4920
               TabIndex        =   166
               Top             =   390
               Width           =   195
            End
            Begin VB.VScrollBar VScroll32 
               Height          =   285
               Left            =   4920
               TabIndex        =   165
               Top             =   870
               Width           =   195
            End
            Begin VB.CommandButton Command7 
               Caption         =   "Calcular"
               Height          =   315
               Left            =   180
               TabIndex        =   164
               Top             =   2520
               Width           =   1365
            End
            Begin VB.Frame Frame15 
               Height          =   615
               Left            =   1860
               TabIndex        =   161
               Top             =   2250
               Width           =   3255
               Begin VB.Label Label34 
                  AutoSize        =   -1  'True
                  Caption         =   "Resultado"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   2010
                  TabIndex        =   163
                  Top             =   240
                  Width           =   930
               End
               Begin VB.Label Label35 
                  AutoSize        =   -1  'True
                  Caption         =   "Resultado"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   150
                  TabIndex        =   162
                  Top             =   240
                  Width           =   930
               End
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   3450
               TabIndex        =   160
               Text            =   "Combo1"
               Top             =   1290
               Width           =   1695
            End
            Begin VB.OptionButton Option8 
               Caption         =   "Milímetros"
               Height          =   285
               Left            =   2280
               TabIndex        =   159
               Top             =   1830
               Width           =   1125
            End
            Begin VB.OptionButton Option9 
               Caption         =   "Pulgadas"
               Height          =   255
               Left            =   4110
               TabIndex        =   158
               Top             =   1830
               Width           =   1005
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               Caption         =   "Peso de corrección (grs)"
               Height          =   195
               Left            =   1890
               TabIndex        =   171
               Top             =   450
               Width           =   1725
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               Caption         =   "Diámetro de mecha (mm)"
               Height          =   195
               Left            =   1860
               TabIndex        =   170
               Top             =   930
               Width           =   1755
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               Caption         =   "Material"
               Height          =   195
               Left            =   3570
               TabIndex        =   169
               Top             =   1350
               Width           =   555
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Unificar Masas"
            Height          =   2895
            Left            =   -74640
            TabIndex        =   136
            Top             =   990
            Width           =   7965
            Begin VB.TextBox Text30 
               Height          =   315
               Left            =   6300
               TabIndex        =   152
               Top             =   1050
               Width           =   1485
            End
            Begin VB.TextBox Text28 
               Height          =   315
               Left            =   4200
               TabIndex        =   151
               Top             =   1050
               Width           =   1485
            End
            Begin VB.TextBox Text29 
               Height          =   315
               Left            =   6300
               TabIndex        =   150
               Top             =   570
               Width           =   1485
            End
            Begin VB.TextBox Text27 
               Height          =   315
               Left            =   4200
               TabIndex        =   149
               Top             =   570
               Width           =   1485
            End
            Begin VB.Frame Frame12 
               Caption         =   "Cantidad de Contrapesos"
               Height          =   1815
               Left            =   150
               TabIndex        =   145
               Top             =   390
               Width           =   2385
               Begin VB.OptionButton Option5 
                  Caption         =   "2 Contrapesos"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   148
                  Top             =   450
                  Width           =   1905
               End
               Begin VB.OptionButton Option6 
                  Caption         =   "3 Contrapesos"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   147
                  Top             =   930
                  Width           =   1665
               End
               Begin VB.OptionButton Option7 
                  Caption         =   "4 Contrapesos"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   146
                  Top             =   1410
                  Width           =   1695
               End
            End
            Begin VB.VScrollBar VScroll27 
               Height          =   285
               Left            =   5460
               TabIndex        =   144
               Top             =   570
               Width           =   195
            End
            Begin VB.VScrollBar VScroll29 
               Height          =   285
               Left            =   7560
               TabIndex        =   143
               Top             =   570
               Width           =   195
            End
            Begin VB.VScrollBar VScroll28 
               Height          =   285
               Left            =   5460
               TabIndex        =   142
               Top             =   1050
               Width           =   195
            End
            Begin VB.VScrollBar VScroll30 
               Height          =   285
               Left            =   7560
               TabIndex        =   141
               Top             =   1050
               Width           =   195
            End
            Begin VB.CommandButton Command6 
               Caption         =   "Calcular"
               Height          =   315
               Left            =   630
               TabIndex        =   140
               Top             =   2340
               Width           =   1365
            End
            Begin VB.Frame Frame13 
               Height          =   615
               Left            =   4200
               TabIndex        =   137
               Top             =   1830
               Width           =   3255
               Begin VB.Label Label30 
                  AutoSize        =   -1  'True
                  Caption         =   "Resultado"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   150
                  TabIndex        =   139
                  Top             =   240
                  Width           =   930
               End
               Begin VB.Label Label31 
                  AutoSize        =   -1  'True
                  Caption         =   "Resultado"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   2010
                  TabIndex        =   138
                  Top             =   240
                  Width           =   930
               End
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "Contrapeso 1"
               Height          =   195
               Left            =   3210
               TabIndex        =   156
               Top             =   630
               Width           =   945
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "Contrapeso 2"
               Height          =   195
               Left            =   3210
               TabIndex        =   155
               Top             =   1110
               Width           =   945
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               Caption         =   "Peso (grs)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   4410
               TabIndex        =   154
               Top             =   300
               Width           =   930
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "Angulo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   6660
               TabIndex        =   153
               Top             =   300
               Width           =   630
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "División del Peso"
            Height          =   3135
            Left            =   -74505
            TabIndex        =   114
            Top             =   855
            Width           =   6015
            Begin VB.TextBox Text25 
               Height          =   315
               Left            =   3270
               TabIndex        =   128
               Top             =   1170
               Width           =   1485
            End
            Begin VB.TextBox Text26 
               Height          =   315
               Left            =   3270
               TabIndex        =   127
               Top             =   1680
               Width           =   1485
            End
            Begin VB.TextBox Text8 
               Height          =   315
               Left            =   4320
               TabIndex        =   126
               Top             =   720
               Width           =   1485
            End
            Begin VB.TextBox Text7 
               Height          =   315
               Left            =   2220
               TabIndex        =   125
               Top             =   720
               Width           =   1485
            End
            Begin VB.VScrollBar VScroll7 
               Height          =   285
               Left            =   3480
               TabIndex        =   124
               Top             =   720
               Width           =   195
            End
            Begin VB.VScrollBar VScroll8 
               Height          =   285
               Left            =   5580
               TabIndex        =   123
               Top             =   720
               Width           =   195
            End
            Begin VB.VScrollBar VScroll25 
               Height          =   285
               Left            =   4530
               TabIndex        =   122
               Top             =   1170
               Width           =   195
            End
            Begin VB.VScrollBar VScroll26 
               Height          =   285
               Left            =   4530
               TabIndex        =   121
               Top             =   1680
               Width           =   195
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Calcular"
               Height          =   345
               Left            =   780
               TabIndex        =   120
               Top             =   2340
               Width           =   1095
            End
            Begin VB.Frame Frame8 
               Height          =   825
               Left            =   3270
               TabIndex        =   115
               Top             =   2070
               Width           =   2505
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Resultado"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   120
                  TabIndex        =   119
                  Top             =   180
                  Width           =   930
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  Caption         =   "Resultado"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   1500
                  TabIndex        =   118
                  Top             =   180
                  Width           =   930
               End
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
                  Caption         =   "Resultado"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   120
                  TabIndex        =   117
                  Top             =   480
                  Width           =   930
               End
               Begin VB.Label Label18 
                  AutoSize        =   -1  'True
                  Caption         =   "Resultado"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   1500
                  TabIndex        =   116
                  Top             =   480
                  Width           =   930
               End
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Peso (grs)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2430
               TabIndex        =   135
               Top             =   450
               Width           =   930
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Angulo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   4500
               TabIndex        =   134
               Top             =   450
               Width           =   630
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Posición 1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2100
               TabIndex        =   133
               Top             =   1230
               Width           =   930
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Posición 2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2100
               TabIndex        =   132
               Top             =   1710
               Width           =   930
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Peso de corrección"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   420
               TabIndex        =   131
               Top             =   750
               Width           =   1755
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Posición 1:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2250
               TabIndex        =   130
               Top             =   2250
               Width           =   975
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Posición 2:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2250
               TabIndex        =   129
               Top             =   2550
               Width           =   975
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Balanceo en 2 planos"
            Height          =   3555
            Left            =   -74730
            TabIndex        =   67
            Top             =   810
            Width           =   10545
            Begin VB.TextBox Text24 
               Height          =   315
               Left            =   7200
               TabIndex        =   105
               Top             =   1650
               Width           =   1485
            End
            Begin VB.VScrollBar VScroll24 
               Height          =   285
               Left            =   8460
               TabIndex        =   104
               Top             =   1650
               Width           =   195
            End
            Begin VB.TextBox Text23 
               Height          =   315
               Left            =   8700
               TabIndex        =   103
               Top             =   1650
               Width           =   1485
            End
            Begin VB.VScrollBar VScroll23 
               Height          =   285
               Left            =   9960
               TabIndex        =   102
               Top             =   1650
               Width           =   195
            End
            Begin VB.TextBox Text22 
               Height          =   315
               Left            =   7200
               TabIndex        =   101
               Top             =   1200
               Width           =   1485
            End
            Begin VB.VScrollBar VScroll22 
               Height          =   285
               Left            =   8460
               TabIndex        =   100
               Top             =   1200
               Width           =   195
            End
            Begin VB.TextBox Text21 
               Height          =   315
               Left            =   8700
               TabIndex        =   99
               Top             =   1200
               Width           =   1485
            End
            Begin VB.VScrollBar VScroll21 
               Height          =   285
               Left            =   9960
               TabIndex        =   98
               Top             =   1200
               Width           =   195
            End
            Begin VB.TextBox Text20 
               Height          =   315
               Left            =   4110
               TabIndex        =   97
               Top             =   1650
               Width           =   1485
            End
            Begin VB.TextBox Text19 
               Height          =   315
               Left            =   5610
               TabIndex        =   96
               Top             =   1650
               Width           =   1485
            End
            Begin VB.TextBox Text17 
               Height          =   315
               Left            =   5610
               TabIndex        =   95
               Top             =   780
               Width           =   1485
            End
            Begin VB.TextBox Text16 
               Height          =   315
               Left            =   4110
               TabIndex        =   94
               Top             =   1200
               Width           =   1485
            End
            Begin VB.TextBox Text15 
               Height          =   315
               Left            =   5610
               TabIndex        =   93
               Top             =   1200
               Width           =   1485
            End
            Begin VB.TextBox Text18 
               Height          =   315
               Left            =   4110
               TabIndex        =   92
               Top             =   780
               Width           =   1485
            End
            Begin VB.TextBox Text13 
               Height          =   315
               Left            =   2010
               TabIndex        =   91
               Top             =   1650
               Width           =   1485
            End
            Begin VB.TextBox Text14 
               Height          =   315
               Left            =   510
               TabIndex        =   90
               Top             =   1650
               Width           =   1485
            End
            Begin VB.TextBox Text9 
               Height          =   315
               Left            =   2010
               TabIndex        =   89
               Top             =   1200
               Width           =   1485
            End
            Begin VB.TextBox Text11 
               Height          =   315
               Left            =   2010
               TabIndex        =   88
               Top             =   780
               Width           =   1485
            End
            Begin VB.TextBox Text10 
               Height          =   315
               Left            =   510
               TabIndex        =   87
               Top             =   1200
               Width           =   1485
            End
            Begin VB.TextBox Text12 
               Height          =   315
               Left            =   510
               TabIndex        =   86
               Top             =   780
               Width           =   1485
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Nueva Corrida de Prueba"
               Height          =   375
               Left            =   2250
               TabIndex        =   85
               Top             =   2880
               Width           =   2205
            End
            Begin VB.Frame Frame9 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Rotación"
               Height          =   525
               Left            =   6930
               TabIndex        =   82
               Top             =   2070
               Width           =   3255
               Begin VB.OptionButton Option3 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Derecho"
                  Height          =   225
                  Left            =   1710
                  TabIndex        =   84
                  Top             =   240
                  Width           =   945
               End
               Begin VB.OptionButton Option4 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Izquierdo"
                  Height          =   225
                  Left            =   540
                  TabIndex        =   83
                  Top             =   240
                  Width           =   945
               End
            End
            Begin VB.Frame Frame10 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Resultados"
               Height          =   525
               Left            =   510
               TabIndex        =   81
               Top             =   2070
               Width           =   6225
            End
            Begin VB.CommandButton Command4 
               Caption         =   "&Calcular"
               Height          =   345
               Left            =   510
               TabIndex        =   80
               Top             =   2910
               Width           =   1485
            End
            Begin VB.VScrollBar VScroll9 
               Height          =   285
               Left            =   3270
               TabIndex        =   79
               Top             =   1200
               Width           =   195
            End
            Begin VB.VScrollBar VScroll10 
               Height          =   285
               Left            =   1770
               TabIndex        =   78
               Top             =   1200
               Width           =   195
            End
            Begin VB.VScrollBar VScroll11 
               Height          =   285
               Left            =   3270
               TabIndex        =   77
               Top             =   780
               Width           =   195
            End
            Begin VB.VScrollBar VScroll12 
               Height          =   285
               Left            =   1770
               TabIndex        =   76
               Top             =   780
               Width           =   195
            End
            Begin VB.VScrollBar VScroll13 
               Height          =   285
               Left            =   3270
               TabIndex        =   75
               Top             =   1650
               Width           =   195
            End
            Begin VB.VScrollBar VScroll14 
               Height          =   285
               Left            =   1770
               TabIndex        =   74
               Top             =   1650
               Width           =   195
            End
            Begin VB.VScrollBar VScroll15 
               Height          =   285
               Left            =   6870
               TabIndex        =   73
               Top             =   1200
               Width           =   195
            End
            Begin VB.VScrollBar VScroll16 
               Height          =   285
               Left            =   5370
               TabIndex        =   72
               Top             =   1200
               Width           =   195
            End
            Begin VB.VScrollBar VScroll17 
               Height          =   285
               Left            =   6870
               TabIndex        =   71
               Top             =   780
               Width           =   195
            End
            Begin VB.VScrollBar VScroll18 
               Height          =   285
               Left            =   5370
               TabIndex        =   70
               Top             =   780
               Width           =   195
            End
            Begin VB.VScrollBar VScroll19 
               Height          =   285
               Left            =   6870
               TabIndex        =   69
               Top             =   1650
               Width           =   195
            End
            Begin VB.VScrollBar VScroll20 
               Height          =   285
               Left            =   5370
               TabIndex        =   68
               Top             =   1650
               Width           =   195
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Angulo"
               Height          =   195
               Left            =   9090
               TabIndex        =   113
               Top             =   990
               Width           =   495
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso (grs)"
               Height          =   195
               Left            =   7500
               TabIndex        =   112
               Top             =   990
               Width           =   705
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso de Prueba"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   7950
               TabIndex        =   111
               Top             =   720
               Width           =   1365
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fase"
               Height          =   195
               Left            =   2520
               TabIndex        =   110
               Top             =   540
               Width           =   345
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vib (mm/S)"
               Height          =   195
               Left            =   840
               TabIndex        =   109
               Top             =   540
               Width           =   780
            End
            Begin VB.Image Image3 
               Height          =   270
               Left            =   150
               Top             =   1230
               Width           =   300
            End
            Begin VB.Image Image4 
               Height          =   270
               Left            =   150
               Top             =   810
               Width           =   300
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Corrida Inicial"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3060
               TabIndex        =   108
               Top             =   300
               Width           =   1185
            End
            Begin VB.Image Image5 
               Height          =   270
               Left            =   150
               Top             =   1650
               Width           =   300
            End
            Begin VB.Image Image6 
               Height          =   270
               Left            =   3750
               Top             =   1230
               Width           =   300
            End
            Begin VB.Image Image7 
               Height          =   270
               Left            =   3750
               Top             =   810
               Width           =   300
            End
            Begin VB.Image Image8 
               Height          =   270
               Left            =   3750
               Top             =   1650
               Width           =   300
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fase"
               Height          =   195
               Left            =   6060
               TabIndex        =   107
               Top             =   540
               Width           =   345
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vib (mm/S)"
               Height          =   195
               Left            =   4380
               TabIndex        =   106
               Top             =   540
               Width           =   780
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Balanceo en 1 plano ventana 1"
            Height          =   3135
            Left            =   -74775
            TabIndex        =   37
            Top             =   765
            Width           =   9105
            Begin VB.Frame Frame5 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Ventanas"
               Height          =   555
               Left            =   270
               TabIndex        =   58
               Top             =   2160
               Width           =   1875
               Begin VB.CheckBox Check1 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "2 Ventanas"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   59
                  Top             =   270
                  Width           =   1665
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Peso de Prueba"
               Height          =   555
               Left            =   5190
               TabIndex        =   56
               Top             =   2160
               Width           =   3255
               Begin VB.CheckBox Check2 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Permance"
                  Height          =   195
                  Left            =   420
                  TabIndex        =   57
                  Top             =   270
                  Width           =   2295
               End
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Nueva Corrida de Prueba"
               Height          =   375
               Left            =   2520
               TabIndex        =   55
               Top             =   2280
               Width           =   2205
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Rotación"
               Height          =   525
               Left            =   5190
               TabIndex        =   52
               Top             =   1620
               Width           =   3255
               Begin VB.OptionButton Option2 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Derecho"
                  Height          =   225
                  Left            =   1710
                  TabIndex        =   54
                  Top             =   240
                  Width           =   945
               End
               Begin VB.OptionButton Option1 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Izquierdo"
                  Height          =   225
                  Left            =   540
                  TabIndex        =   53
                  Top             =   240
                  Width           =   945
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Resultados"
               Height          =   525
               Left            =   2400
               TabIndex        =   51
               Top             =   1620
               Width           =   2745
            End
            Begin VB.CommandButton Command1 
               Caption         =   "&Calcular"
               Height          =   405
               Left            =   180
               TabIndex        =   50
               Top             =   1680
               Width           =   1485
            End
            Begin VB.VScrollBar VScroll6 
               Height          =   285
               Left            =   8250
               TabIndex        =   49
               Top             =   1140
               Width           =   195
            End
            Begin VB.TextBox Text6 
               Height          =   315
               Left            =   6990
               TabIndex        =   48
               Top             =   1140
               Width           =   1485
            End
            Begin VB.VScrollBar VScroll5 
               Height          =   285
               Left            =   6750
               TabIndex        =   47
               Top             =   1140
               Width           =   195
            End
            Begin VB.TextBox Text5 
               Height          =   315
               Left            =   5490
               TabIndex        =   46
               Top             =   1140
               Width           =   1485
            End
            Begin VB.VScrollBar VScroll4 
               Height          =   285
               Left            =   5160
               TabIndex        =   45
               Top             =   1140
               Width           =   195
            End
            Begin VB.TextBox Text4 
               Height          =   315
               Left            =   3900
               TabIndex        =   44
               Top             =   1140
               Width           =   1485
            End
            Begin VB.VScrollBar VScroll3 
               Height          =   285
               Left            =   3660
               TabIndex        =   43
               Top             =   1140
               Width           =   195
            End
            Begin VB.TextBox Text3 
               Height          =   315
               Left            =   2400
               TabIndex        =   42
               Top             =   1140
               Width           =   1485
            End
            Begin VB.VScrollBar VScroll2 
               Height          =   285
               Left            =   5160
               TabIndex        =   41
               Top             =   720
               Width           =   195
            End
            Begin VB.TextBox Text2 
               Height          =   315
               Left            =   3900
               TabIndex        =   40
               Top             =   720
               Width           =   1485
            End
            Begin VB.VScrollBar VScroll1 
               Height          =   285
               Left            =   3660
               TabIndex        =   39
               Top             =   720
               Width           =   195
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Left            =   2400
               TabIndex        =   38
               Top             =   720
               Width           =   1485
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Angulo"
               Height          =   195
               Left            =   7200
               TabIndex        =   66
               Top             =   900
               Width           =   495
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso (grs)"
               Height          =   195
               Left            =   5760
               TabIndex        =   65
               Top             =   900
               Width           =   705
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso de Prueba"
               Height          =   195
               Left            =   5880
               TabIndex        =   64
               Top             =   570
               Width           =   1140
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fase"
               Height          =   195
               Left            =   4320
               TabIndex        =   63
               Top             =   450
               Width           =   345
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vib (mm/S)"
               Height          =   195
               Left            =   2670
               TabIndex        =   62
               Top             =   420
               Width           =   780
            End
            Begin VB.Image Image2 
               Height          =   270
               Left            =   2040
               Top             =   1170
               Width           =   300
            End
            Begin VB.Image Image1 
               Height          =   270
               Left            =   2040
               Top             =   750
               Width           =   300
            End
            Begin VB.Label lblCorrPrueba 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Corrida de Prueba"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   420
               TabIndex        =   61
               Top             =   1200
               Width           =   1545
            End
            Begin VB.Label lblCorrInicial 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Corrida Inicial"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   780
               TabIndex        =   60
               Top             =   780
               Width           =   1185
            End
         End
         Begin VB.Image Image9 
            Height          =   7095
            Left            =   -75000
            Top             =   585
            Width           =   9975
         End
      End
      Begin VB.PictureBox Scope 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2505
         Index           =   1
         Left            =   180
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   688
         TabIndex        =   35
         Top             =   2625
         Width           =   10320
      End
      Begin VB.PictureBox Scope 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   0
         Left            =   180
         ScaleHeight     =   121
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   686
         TabIndex        =   34
         Top             =   765
         Width           =   10290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gráfico de Tendencias del Balanceo Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -68610
         TabIndex        =   232
         Top             =   540
         Width           =   4470
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   17610
      TabIndex        =   30
      Top             =   3750
      Width           =   705
   End
   Begin VB.HScrollBar Angl 
      Height          =   315
      Left            =   6450
      Max             =   360
      TabIndex        =   29
      Top             =   5400
      Width           =   2475
   End
   Begin VB.Timer tmrLINE 
      Interval        =   1000
      Left            =   8700
      Top             =   150
   End
   Begin VB.Frame Frame24 
      Caption         =   "Filtro Pasa Bajos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   9030
      TabIndex        =   20
      Top             =   390
      Width           =   7335
      Begin VB.PictureBox picPasaBajos 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   30
         ScaleHeight     =   71
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   359
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   5415
      End
      Begin VB.PictureBox picLPKernel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   5460
         ScaleHeight     =   71
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   119
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   1815
      End
      Begin VB.PictureBox picSpecLP 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1000
         Left            =   30
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   359
         TabIndex        =   22
         Top             =   1530
         Width           =   5415
      End
      Begin VB.PictureBox picSpecLPK 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1000
         Left            =   5460
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   119
         TabIndex        =   21
         Top             =   1530
         Width           =   1815
      End
      Begin VB.Label lblKernel 
         AutoSize        =   -1  'True
         Caption         =   "Núcleo de Filtro"
         Height          =   195
         Left            =   5670
         TabIndex        =   26
         Top             =   30
         Width           =   1110
      End
      Begin VB.Label lblFreqSpec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espectro de Frecuencia"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1695
      End
   End
   Begin VB.Frame Frame23 
      Caption         =   "Filtro Pasa Altos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   9030
      TabIndex        =   13
      Top             =   2970
      Width           =   7335
      Begin VB.PictureBox picPasaAltos 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   30
         ScaleHeight     =   71
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   359
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   5415
      End
      Begin VB.PictureBox picHPKernel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   5460
         ScaleHeight     =   71
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   119
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   1815
      End
      Begin VB.PictureBox picSpecHP 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1000
         Left            =   30
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   359
         TabIndex        =   15
         Top             =   1530
         Width           =   5415
      End
      Begin VB.PictureBox picSpecHPK 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1000
         Left            =   5460
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   119
         TabIndex        =   14
         Top             =   1530
         Width           =   1815
      End
      Begin VB.Label lblKernel2 
         AutoSize        =   -1  'True
         Caption         =   "Núcleo de Filtro"
         Height          =   195
         Left            =   5670
         TabIndex        =   19
         Top             =   0
         Width           =   1110
      End
      Begin VB.Label lblFreqSpec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espectro de Frecuencia"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1695
      End
   End
   Begin VB.VScrollBar scrlFactor 
      Height          =   3015
      LargeChange     =   20
      Left            =   16770
      Max             =   1000
      SmallChange     =   10
      TabIndex        =   10
      Top             =   1980
      Value           =   880
      Width           =   315
   End
   Begin VB.VScrollBar scrlTaps 
      Height          =   3015
      LargeChange     =   10
      Left            =   17160
      Max             =   512
      Min             =   2
      TabIndex        =   9
      Top             =   1980
      Value           =   450
      Width           =   315
   End
   Begin VB.Frame Frame22 
      Caption         =   "Tiempos:"
      Height          =   1275
      Left            =   16380
      TabIndex        =   5
      Top             =   420
      Width           =   1665
      Begin VB.Label lblTimeFilter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro Avg.:"
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   300
         Width           =   750
      End
      Begin VB.Label lblTimeFFT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FFT Avg.:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   600
         Width           =   705
      End
      Begin VB.Label lblTimeTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   900
         Width           =   405
      End
   End
   Begin VB.Frame Stuff 
      Height          =   510
      Left            =   30
      TabIndex        =   1
      Top             =   5250
      Width           =   3360
      Begin VB.CommandButton StartButton 
         Caption         =   "&Iniciar"
         Height          =   336
         Left            =   150
         TabIndex        =   4
         Top             =   120
         Width           =   804
      End
      Begin VB.CommandButton StopButton 
         Caption         =   "&Parar"
         Enabled         =   0   'False
         Height          =   336
         Left            =   1140
         TabIndex        =   3
         Top             =   120
         Width           =   804
      End
      Begin VB.CheckBox Flicker 
         Caption         =   "Actualizar"
         Height          =   300
         Left            =   2190
         TabIndex        =   2
         Top             =   150
         Width           =   1035
      End
   End
   Begin VB.ComboBox DevicesBox 
      Height          =   315
      Left            =   72
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   72
      Width           =   3108
   End
   Begin VB.Label lblS 
      Caption         =   "Label55"
      Height          =   285
      Left            =   5310
      TabIndex        =   231
      Top             =   5445
      Width           =   915
   End
   Begin VB.Label lblF 
      Caption         =   "Label50"
      Height          =   285
      Left            =   5310
      TabIndex        =   230
      Top             =   5085
      Width           =   915
   End
   Begin VB.Label lblFFT 
      Caption         =   "Label50"
      Height          =   195
      Left            =   9135
      TabIndex        =   228
      Top             =   5580
      Width           =   2085
   End
   Begin VB.Label lblFrecMac 
      Caption         =   "Frec.Max"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   6780
      TabIndex        =   32
      Top             =   5070
      Width           =   1665
   End
   Begin VB.Label lblSeñal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblSeñal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   16770
      TabIndex        =   31
      Top             =   5340
      Width           =   1125
   End
   Begin VB.Label lblY 
      Caption         =   "lblY"
      Height          =   285
      Left            =   3465
      TabIndex        =   28
      Top             =   5445
      Width           =   1155
   End
   Begin VB.Label lblX 
      Caption         =   "lblX"
      Height          =   285
      Left            =   3510
      TabIndex        =   27
      Top             =   5085
      Width           =   1035
   End
   Begin VB.Label lblTaps 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Corte (100):"
      Height          =   195
      Left            =   17160
      TabIndex        =   12
      Top             =   5010
      Width           =   825
   End
   Begin VB.Label lblFactor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Factor (0.2):"
      Height          =   195
      Left            =   16740
      TabIndex        =   11
      Top             =   1770
      Width           =   855
   End
End
Attribute VB_Name = "frmBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const PI As Double = 3.14159265358979
Private Señal_Entrada(Muestras - 1) As Single

Private Type PointAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Dim Pt As PointAPI
Dim ScreenStartX As Long, ScreenEndX As Long
Dim imgStartX As Long, imgEndX As Long
Dim dX As Long, imgLeft As Long
Dim StartMove As Byte
Dim LimitMin As Long, LimitMax As Long, CurrentPos As Long
Dim sW As Long, sH As Long, CenX As Long, CenY As Long
Dim TempKonstant As Single, SpeedKonstant As Single
'Const PI As Single = 3.14159265358979
Const PIBY180 As Single = PI / 180
Dim XL As Long, yL As Long, PlotX As Long, PlotY As Long, MrySize As Double, Ventana As Double
'------------------------------------------------------------------------------------------
Dim maxvol As Long, Hz As Long, oscila As Long
Dim HzColor As Long, xMax As Integer, HzTip As Long
'------------------------------------------------------------------------------------------
'Dim InData(0 To Muestras - 1) As Integer
'Dim OutData(0 To Muestras - 1) As Single
Dim SeñalEntrada As Single
Dim cx, cx1 As Integer
Dim cy, cy1 As Integer
Dim X As Single
Dim Y As Single
Dim radio As Single

Private ThisAngle           As Single
Private Const Radius        As Long = 45



Private Sub SenXY(H As Single)
'(X-K)^2+(Y-H)^2=(R^2) ->definicion de un circulo
'(2250-K)^2+(2100-H)^2=(1848.5)^2->tambie es = 3416948 en la ecuacion original
'picPolar.Cls  '
'Dim H As Single
'VS1.Value = cy1 - 500
'VH1.Value = cx1 - 500
'H = VS1.Value
radio = VH1.value
                          '(radio)
X = ((Cos(H * 3.141 / 1000) * radio)) + cx1  '2000
Y = ((Sin(H * 3.141 / 1000) * radio)) + cy1  '2000
'picPolar.Line (cx1, cy1)-(X, cy - Y), 0 'para la graficación "X" en la segunda parte no hay que agregarle "cx"
End Sub

Private Sub cmdCerrar_Click()
Call FinDispositivoSeñal
    Unload Me
End Sub

Private Sub Flicker_Click()
   'On Error GoTo Err_Proc
    Scope(0).Cls
    Scope(1).Cls
    If Flicker.value = vbChecked Then
        Scope(0).AutoRedraw = True
        Scope(1).AutoRedraw = True
    Else
        Scope(0).AutoRedraw = False
        Scope(1).AutoRedraw = False
    End If
Exit_Proc:
   Exit Sub
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "Base", "Flicker_Click"
   Err.Clear
   Resume Exit_Proc
End Sub

'****************************************
'Rutina principal que carga el formulario
Private Sub Form_Load()
Dim I As Integer
    'cy = 'picPolar.ScaleHeight
    'cx = 'picPolar.ScaleWidth
    
    'cy1 = 'picPolar.ScaleHeight \ 2
    'cx1 = 'picPolar.ScaleWidth \ 2
'On Error GoTo Err_Proc
    Call IniciarSonido(Me)  'Llene el DevicesBox
    Call DoReverse   'Pre-cálculo de estos

InitFFT                     ' preparar tablas de referencia para la FFT
InitFastConvolution         ' utilizar el algoritmo de convolución escrito en C
grpTendencias.MaxValue = 1000
CargarGrpTendencias_Click

CenX = sW \ 2: CenY = sH \ 2
LimitMin = 0: LimitMax = 360

'picPIzq.ScaleMode = 0
'Establecer MinWidth y MinHeight basada en la forma ...
Dim XAjuste As Long, YAjuste As Long
XAjuste = Me.Width \ Screen.TwipsPerPixelX - Me.ScaleWidth
YAjuste = Me.Height \ Screen.TwipsPerPixelY - Me.ScaleHeight

ShapeCtrl Scope(0), 10, 0
ShapeCtrl Scope(1), 10, 0

'picPDer.ScaleMode = 3
Exit_Proc:
   Exit Sub
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "Base", "Form_Load"
   Err.Clear
   Resume Exit_Proc
End Sub

' Pegar en la sección Declaraciones de Form1.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim Msg   ' Declara la variable.
   If UnloadMode > 0 Then
      Msg = "Esta a punto de cerrar el balanceo/análisis actual, desea continuar?"
   End If
   ' Si el usuario hace clic en el botón No, se detiene QueryUnload.
   If MsgBox(Msg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then
        Cancel = True
   Else
        If DevHandle <> 0 Then
            Call FinDispositivoSeñal
        End If
   End If
End Sub

'***********************************************************************************************
'Redimensiona el formulario y los controles ubucados en el formulario
Private Sub Form_Resize()
   'On Error GoTo Err_Proc
    Scope(0).Cls
    Scope(1).Cls
    DevicesBox.Width = Me.ScaleWidth - 13
    Scope(0).ScaleHeight = 256
    Scope(0).ScaleWidth = 255
    Scope(1).ScaleHeight = 256
    Scope(1).ScaleWidth = 255
    'Hacer que el tamaño de la ventana ahora para que no interfiera con el nuevo trazado de los datos
    DoEvents
    'Volver a dibujar los datos en el nuevo tamaño
    If Inited = True Then
        Call DibujarSeñal
    End If
Exit_Proc:
   Exit Sub
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "Base", "Form_Resize"
   Err.Clear
   Resume Exit_Proc
End Sub

Private Sub CargarGrpTendencias_Click()
Dim I As Integer
Dim tval As Single
'add some random data
For I = 1 To HScroll1.value
    tval = Int(Rnd * 999)
    'actual altitude
    grpTendencias.Values.Add tval
    'tooltip description of blip
    grpTendencias.aDescription.Add tval
Next I
'force redraw
grpTendencias.ReDraw
End Sub

Private Sub HScroll1_Change()
Dim I As Integer
Dim tval As Single
For I = 1 To HScroll1.value
    tval = Int(Rnd * 999)
    grpTendencias.Values.Add tval
    grpTendencias.aDescription.Add tval
Next I
grpTendencias.ReDraw
End Sub

Private Sub StartButton_Click()
    Static WAVEFORMAT As WAVEFORMATEX
    'On Error GoTo Err_Proc

    With WAVEFORMAT
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 2 'Dos Canales - izquierdo y derecho
        .SamplesPerSec = 44100 '44.1Khz   '11025 '11khz
        .BitsPerSample = 8
        .BLOCKALIGN = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BLOCKALIGN * .SamplesPerSec
        .ExtraDataSize = 0
    End With
        maxvol = waveInOpen(DevHandle, DevicesBox.ListIndex, VarPtr(WAVEFORMAT), 0, 0, 0)

    Debug.Print "waveInOpen:"; waveInOpen(DevHandle, DevicesBox.ListIndex, VarPtr(WAVEFORMAT), 0, 0, 0)
    If DevHandle = 0 Then
        Call MsgBox("Imposible abrir el dispositivo de entrada de señal!", vbExclamation, "Ooops!")
        'Exit Sub
    End If
    Debug.Print " "; DevHandle
    Call waveInStart(DevHandle)
    
    Inited = True
       
    StopButton.Enabled = True
    StartButton.Enabled = False
    
    Call Visualize
Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "Base", "StartButton_Click"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub StopButton_Click()
    Call FinDispositivoSeñal
End Sub

Private Sub FinDispositivoSeñal()
   'On Error GoTo Err_Proc
    Call waveInReset(DevHandle)
    Call waveInClose(DevHandle)
    DevHandle = 0
    StopButton.Enabled = False
    StartButton.Enabled = True
Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "Base", "FinDispositivoSeñal"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub scrlFactor_Change()
   ' ActualizarDisplay
End Sub
Private Sub scrlFactor_Scroll()
    'ActualizarDisplay
End Sub
Private Sub scrlTaps_Change()
    'ActualizarDisplay
End Sub
Private Sub scrlTaps_Scroll()
    'ActualizarDisplay
End Sub
Private Sub ActualizarInpDisplay()
    'picSeñalEntrada.Cls
    'picSpecFrecuencia.Cls
    'Mostrar picSeñalEntrada, Señal_Entrada, True, True
    'MostrarFFT picSpecFrecuencia, Señal_Entrada, True, True    'False
End Sub
Private Sub Scope_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If X < Muestras \ 2 - 1 Then
   lblY.Caption = " " & CLng(44100 / Muestras * X) & " Hz"
End If

End Sub
                                            
Private Sub DibujarRadio(pic As PictureBox, ByVal radio As Single, ByVal Angulo As Single, Color)
pic.Cls
Dim xp, yp, rx, ry, cx, cy, rxg, ryg As Long
'PI = 4 * Atn(1)
cx = pic.ScaleWidth / 2     'pic.CurrentX
cy = pic.ScaleHeight / 2    'pic.CurrentY
'Angulo en Grados
Angulo = Angulo Mod 360
Angulo = Angulo * PI / 180
xp = 0
yp = Abs(radio)
rx = xp * Cos(Angulo) - yp * Sin(Angulo)
ry = xp * Sin(Angulo) + yp * Cos(Angulo)
rxg = cx + rx
ryg = cy - ry
pic.Line (cx, cy)-(rxg, ryg), Color
DoEvents
End Sub

Private Sub ActualizarDisplay()
    Dim SeñalFactor   As Single
    Dim lngTaps     As Long
   'On Error GoTo Err_Proc
    SeñalFactor = scrlFactor.value / scrlFactor.Max * 0.5
    lngTaps = scrlTaps.value
    lblFactor.Caption = "Factor (" & Format(SeñalFactor, "0.00") & "):"
    lblTaps.Caption = "corte (" & lngTaps & "):"
    Filtro lngTaps, SeñalFactor
        DibujarRadio picPDer, 500, CSng(SeñalFactor), vbRed

Exit_Proc:
   Exit Sub
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPloteaSeñal", "ActualizarDisplay"
   Err.Clear
   Resume Exit_Proc
End Sub
Private Sub Visualize()
    Static X As Long
    Static Wave As WAVEHDR
    Static InData(0 To NumSamples - 1) As Integer
    Static OutData(0 To NumSamples - 1) As Single

    Wave.lpData = VarPtr(InData(0))
    Wave.dwBufferLength = 1024 'Esto ahora se siembran 512 Todavía hay 256 muestras por canal
    Wave.dwFlags = 0

    Do
        Call waveInPrepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
        Call waveInAddBuffer(DevHandle, VarPtr(Wave), Len(Wave))
        maxvol = Wave.lpData   '        waveInOpen(DevHandle, DevicesBox.ListIndex, VarPtr(WAVEFORMAT), 0, 0, 0)
        Do  'Nada - esperar a que el controlador de audio muestree las dos señales de la onda.
        Loop Until ((Wave.dwFlags And WHDR_DONE) = WHDR_DONE) Or DevHandle = 0

        If DevHandle = 0 Then Exit Do 'El dispositivo se ha cerrado...
        Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
        Call FFTAudio(InData, OutData)
        Scope(0).Cls
        Scope(1).Cls
        '------------------------------------------------------------------------------------------------------------------
        Dim c As Double, LowMidHig
        For X = 1 To 511
            'ScopeBuff.PSet (X, ScopeHeight - OutData(X) / 500), oscila ' FFT out (just for reference)
            Señal_Entrada(X) = CSng(OutData(X * 2))
            SeñalEntrada = CSng(InData(X * 2))
            'If Abs(InData(X)) > maxvol Then
            maxvol = Abs(OutData(X))
            Hz = Int(44100 * X) / 1024
            'Label1.Caption = Hz & " Hz"
            lblF.Caption = Hz
            'HzColor = vbRed '+ Hz
            'LowMidHig = ScopeHeight
            xMax = X
            'End If
        Next
        X = xMax
        lblF = X
        'Señal_Entrada = InData
        c = 0.5 * (1 - Cos(X * 2 * 3.1416 / 512)) 'Hanning Window
        OutData(X) = c * OutData(X)
        maxvol = 0
        Call DibujarSeñal
        DoEvents
        '------------------------------------------------------------------------------------------------------------------
    Loop While DevHandle <> 0 'Mientras que el dispositivo de audio está abierto
Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "Base", "Visualize"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub DibujarSeñal()
    Dim a, b, I As Integer
    Static X As Long
   'On Error GoTo Err_Proc
    Scope(0).CurrentX = -1
    Scope(0).CurrentY = Scope(0).ScaleHeight \ 2
    Scope(1).CurrentX = -1
    Scope(1).CurrentY = Scope(0).ScaleHeight \ 2
    'Plotear los datos...
    '-----------------------------------------------------------------------------------------------
    Mostrar Scope(0), Señal_Entrada, True, True
    MostrarFFT Scope(1), Señal_Entrada, True, True    'False
    DibujarRadio picPIzq, 100, SeñalEntrada, vbRed
    DibujarRadio picPDer, 500, CSng(SeñalEntrada), vbRed
    'mAngulo Val(Hz), 400, 'picPolar1
    DoEvents
    'Call SenXY(SeñalEntrada)
    'DibujarLinea 20, Hz, picPIzq
    ActualizarDisplay
    Scope(0).CurrentY = Scope(0).Width
    Scope(1).CurrentY = Scope(0).Width
    
Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "Base", "DibujarSeñal"
   Err.Clear
   Resume Exit_Proc
End Sub

' Fourier de transformación de datos y el gráfico de resultados
Private Sub MostrarFFT(pb As PictureBox, data() As Single, ByVal window As Boolean, ByVal Normalizar As Boolean)
    Dim EntradaSeñalReal(Muestras - 1)       As Single
    Dim SalidaSeñalReal(Muestras - 1)        As Single
    Dim SalidaSeñalImaginaria(Muestras - 1)  As Single
    Dim SalidaSeñalComp(Muestras / 2 - 1)    As Single
    Dim I                                    As Long
    Dim DivisorSeñal                         As Single
    ' Hamming ventana por menos fugas
   ''On Error GoTo Err_Proc
    For I = 0 To UBound(data)
        If window Then
            EntradaSeñalReal(I) = data(I) * VentanaHamming(I, UBound(data) + 1)
        Else
            EntradaSeñalReal(I) = data(I)
        End If
    Next
    ' transformar los datos en el dominio de las frecuencias
    RealFFT Muestras, EntradaSeñalReal, SalidaSeñalReal, SalidaSeñalImaginaria
    ' magnitud del espectro a escala del complejo transformado
    DivisorSeñal = Muestras / 8
    For I = 0 To Muestras / 2 - 1
        SalidaSeñalComp(I) = Sqr(SalidaSeñalReal(I) * SalidaSeñalReal(I) + SalidaSeñalImaginaria(I) * SalidaSeñalImaginaria(I)) / DivisorSeñal
    Next
    Mostrar pb, SalidaSeñalComp, False, Normalizar
    DoEvents
Exit_Proc:
   Exit Sub
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmBase", "MostrarFFT"
   Err.Clear
   Resume Exit_Proc
End Sub

' representar una función / de imagen a escala en un PictureBox
Sub Mostrar(pb As PictureBox, SeñalMuestras() As Single, ByVal Centrar As Boolean, ByVal Normalizar As Boolean)
    Dim dy        As Long, dy2          As Long
    Dim n         As Long, I            As Long
    Dim SeñalMax  As Single, SeñalVal   As Single
    Dim yL        As Single, yN         As Single
    Dim X         As Single, k          As Single
    Dim st        As Single
   ''On Error GoTo Err_Proc
    dy = pb.ScaleHeight - 1
    dy2 = dy \ 2
    n = UBound(SeñalMuestras) + 1
    st = n / pb.ScaleWidth
    If Normalizar Then
        For I = 0 To n - 1
            If Abs(SeñalMuestras(I)) > SeñalMax Then SeñalMax = Abs(SeñalMuestras(I))
        Next
    End If
    If SeñalMax = 0 Then SeñalMax = 1
    SeñalVal = SeñalMuestras(0) / SeñalMax
    If Centrar Then    'para centrar la señal en el picture
        yL = -SeñalVal * dy2 + dy2
        pb.ForeColor = vbBlack    'RGB(160, 160, 180)
        pb.Line (0, dy2)-(pb.ScaleWidth, dy2)
        pb.ForeColor = vbBlue   'vbBlack
    Else
        yL = dy - SeñalVal * dy
    End If
    k = k + st
    Do
        SeñalVal = SeñalMuestras(Fix(k)) / SeñalMax
        
        If Centrar Then  'para quitar el centrado de la señal en el picture
            yN = -SeñalVal * dy2 + dy2
        Else
            yN = dy - SeñalVal * dy
        End If
        pb.Line (X, yL)-(X + 1, yN), vbBlack
        'pb.Line -(X, yN), vbGreen
        yL = yN
        X = X + 1
        k = k + st
        
      'Picture1.Line (X, Y)-(i, j), ((&HFFFFFF / (i + 1)))
      'X = i
      'Y = j
      'i = i + 1
      'Debug.Print "i= " & i
      
    Loop While k < n
DoEvents
Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPloteaSeñal", "Mostrar"
   Err.Clear
   Resume Exit_Proc

End Sub

' filtrar una señal y plotear el resultado
Private Sub Filtro(ByVal lngTaps As Long, ByVal SeñalFactor As Single)
    Dim sngLP()         As Single
    Dim sngHP()         As Single
    Dim udtLP           As FilterKernel
    Dim udtHP           As FilterKernel
    Dim D               As Double
    Dim tmrTotal        As Double
    Dim tmrFilter       As Double
    Dim tmrFFT          As Double
    
   'On Error GoTo Err_Proc

    tmrTotal = Timer
    
    udtLP = CrearFiltro(FiltroPasaBajos, lngTaps, SeñalFactor)
    udtHP = CrearFiltro(FiltroPasaAltos, lngTaps, SeñalFactor)
    
    D = Timer
    sngLP = Señal_Entrada
    FiltrarProceso sngLP, udtLP
        picPasaBajos.Cls: Mostrar picPasaBajos, sngLP, True, True
        picLPKernel.Cls:  Mostrar picLPKernel, udtLP.kernel, True, False
    tmrFilter = Timer - D

    D = Timer
    picSpecLP.Cls:      MostrarFFT picSpecLP, sngLP, True, True
    picSpecLPK.Cls:     MostrarFFT picSpecLPK, udtLP.kernel, False, True
    tmrFFT = Timer - D
    
    D = Timer
    sngHP = Señal_Entrada
    FiltrarProceso sngHP, udtHP
    picPasaAltos.Cls:    Mostrar picPasaAltos, sngHP, False, True
    picHPKernel.Cls:    Mostrar picHPKernel, udtHP.kernel, True, False
    tmrFilter = (tmrFilter + (Timer - D)) / 2

    D = Timer
    picSpecHP.Cls:      MostrarFFT picSpecHP, sngHP, True, True
    picSpecHPK.Cls:     MostrarFFT picSpecHPK, udtHP.kernel, False, True
    tmrFFT = (tmrFFT + (Timer - D)) / 4
    
    lblTimeFilter.Caption = "Filtro Avg.: " & Round(tmrFilter * 1000) & " ms"
    lblTimeFFT.Caption = "FFT Avg.: " & Round(tmrFFT * 1000) & " ms"
    lblTimeTotal.Caption = "Total: " & Round((Timer - tmrTotal) * 1000) & " ms"


Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPloteaSeñal", "Filtro"
   Err.Clear
   Resume Exit_Proc

End Sub

Function FormatComplex(NumToFormat As String, RealFormatCode As String, ImagFormatCode As String)
    Dim PlusOrMinus As String
    Dim CharPosition As Integer

    ' Is NumToFormat real?
   'On Error GoTo Err_Proc

    If Right(NumToFormat, 1) <> "i" Then
        ' NumToFormat is real.
        FormatComplex = Format(NumToFormat, RealFormatCode)
    Else
        ' NumToFormat is either imaginary or complex.
        ' Search NumToFormat from right until + or - or left end is
        ' reached.
        PlusOrMinus = "not found"
        For CharPosition = Len(NumToFormat) - 1 To 1 Step -1
            PlusOrMinus = Mid(NumToFormat, CharPosition, 1)
            If PlusOrMinus = "+" Or PlusOrMinus = "-" Then Exit For
        Next
        ' Is NumToFormat complex or imaginary?
        If (PlusOrMinus = "+" Or PlusOrMinus = "-") And _
            CharPosition <> 1 Then
            ' NumToFormat is complex.
            ' Is imaginary component negative?
            If Mid(NumToFormat, CharPosition, _
                Len(NumToFormat) - CharPosition) < 0 Then
                ' Imaginary component is negative, so "-" does not need
                ' to be added.
                FormatComplex = Format(Left(NumToFormat, _
                    CharPosition - 1), RealFormatCode) & _
                    Format(Mid(NumToFormat, CharPosition, _
                    Len(NumToFormat) - CharPosition), _
                    ImagFormatCode) & "i"
            Else
                ' Imaginary component is not negative, so "+" needs to
                ' be added.
                FormatComplex = Format(Left(NumToFormat, _
                    CharPosition - 1), RealFormatCode) & "+" & _
                    Format(Mid(NumToFormat, CharPosition, _
                    Len(NumToFormat) - CharPosition), _
                    ImagFormatCode) & "i"
            End If
        Else
            ' NumToFormat is imaginary.
            FormatComplex = Format(Left(NumToFormat, Len(NumToFormat) - 1), ImagFormatCode) & "i"
        End If
    End If

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmPloteaSeñal", "FormatComplex"
   Err.Clear
   Resume Exit_Proc

End Function

'**************************************************************************************************************************
'Mediante esta subrutina se puede encontrar facilmente el
'máximo valor de una matriz cualquiera
Public Function FindMax(ByRef Matriz() As Double, Optional a As Long) As Double
Dim Max As Double, I As Integer
Max = Matriz(LBound(Matriz) + a)
For I = LBound(Matriz) + a To MrySize \ Ventana
    If Matriz(I) > Max Then
        Max = Matriz(I)
    End If
Next I
FindMax = Max
End Function
'Mediante esta subrutina se puede encontrar facilmente el
'míniimo valor de una matriz cualquiera
Public Function FindMin(ByRef Matriz() As Double) As Double
Dim Min As Double, I As Integer
Min = Matriz(LBound(Matriz))
For I = LBound(Matriz) To MrySize \ Ventana
    If Matriz(I) < Min Then
        Min = Matriz(I)
    End If
Next I
FindMin = Min
End Function

Private Sub DibujarLinea(ByVal largo As Single, ByVal Angulo As Single, pic As PictureBox)
Dim cx_i, cy_i, xp, yp, rx, ry, rxg, ryg, grados, CenterX As Long
'Dim PI As Single

'PI = 4 * Atn(1)
cx_i = pic.Width / 2  ' CurrentX
cy_i = pic.Height / 2  'CurrentY
'Angulo en Grados
Angulo = Angulo Mod 360
Angulo = Angulo * PI / 180
' lblFFT = Angulo
xp = 0
yp = Abs(largo)
rx = xp * Cos(Angulo) - yp * Sin(Angulo)
ry = xp * Sin(Angulo) + yp * Cos(Angulo)
rxg = cx_i + rx
ryg = cy_i - ry

''picPolar.Line (cx_i, cy_i)-(rxg, ryg)
''picPolar.Line (cx_i, cy_i)-(rxg, ryg)

''picPolar.Circle (rxg, ryg), 2, vbBlue
''picPolar.Circle (rxg, ryg), 3, vbGreen
If Sqr((rx - CenterX) ^ 2 + (ry - Y) ^ 2) >= Radius Then 'not too close to center
    ThisAngle = Atn2(rxg - cx_i, cy_i - ryg) 'y increases towards the bottom
    lblFFT = Format$(ThisAngle, "000.0°")
End If
' si la longitud negativa volver a la posición inicial
If largo < 0 Then
    pic.CurrentX = cx_i
    pic.CurrentY = cy_i
End If
End Sub

Private Function Atn2(ByVal X As Single, ByVal Y As Single) As Single
Dim TwoPi
  'computes the angle in degrees from (relative) mouse coords
  'quadrants are numbered counterclockwise(!) as follows; the o indicating the center
  '            |
  '            |
  '     II     |     I
  '            |
  '            |
  ' -----------o-----------
  '            |
  '            |
  '    III     |     IV
  '            |
  '            |
    If X = 0 Then
        X = 1E-16  'prevent infinity
    End If
    Atn2 = Atn(Abs(Y) / Abs(X)) 'returns the correct value for quadrant 1 only
    Select Case True
      Case X < 0 And Y >= 0     'quadrant II
        Atn2 = PI - Atn2        'adjust for q2
      Case X < 0 And Y < 0      'quadrant III
        Atn2 = Atn2 + PI        'adjust for q3
      Case X >= 0 And Y < 0     'quadrant IV
        Atn2 = TwoPi - Atn2     'adjust for q4
    End Select
    Atn2 = Atn2 * 180 / PI      'convert to degrees
End Function

Sub mAngulo(Frecuencia&, Amplitud&, pic As PictureBox)
Dim X1, X2, Y1, Y2, grados As Single
'Dim Pi As Double
'Pi = 3.14159265358979
X1 = Int(pic.Width / 2)
Y1 = Int(pic.Height / 2)
' Pic1.Cls
grados = (Round(Frecuencia, 2) * PI / 180) - (PI / 2)  'Acimut
pic.Line (X1, Y1)-Step(Amplitud * Cos(grados), Amplitud * Sin(grados)), vbBlack

pic.Circle (X1 + Amplitud * Cos(grados), Y1 + Amplitud * Sin(grados)), 2, vbBlue
pic.Print Int(Cos(grados) * PI / 180)
End Sub
