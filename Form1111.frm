VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFields 
      DataField       =   "ID_Rodam"
      Height          =   285
      Index           =   0
      Left            =   2685
      TabIndex        =   13
      Top             =   630
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Designacion"
      Height          =   285
      Index           =   1
      Left            =   2685
      TabIndex        =   12
      Top             =   945
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Marca_Rodam"
      Height          =   285
      Index           =   2
      Left            =   2685
      TabIndex        =   11
      Top             =   1290
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "d_Diam_Interno"
      Height          =   285
      Index           =   3
      Left            =   2685
      TabIndex        =   10
      Top             =   1605
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "D_Diam_Externo"
      Height          =   285
      Index           =   4
      Left            =   2685
      TabIndex        =   9
      Top             =   1935
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "B_Ancho"
      Height          =   285
      Index           =   5
      Left            =   2685
      TabIndex        =   8
      Top             =   2250
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "a_Ang_Contacto"
      Height          =   285
      Index           =   6
      Left            =   2685
      TabIndex        =   7
      Top             =   2565
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "fa_Factor"
      Height          =   285
      Index           =   7
      Left            =   2685
      TabIndex        =   6
      Top             =   2895
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C_Cap_Carga_Dinamica"
      Height          =   285
      Index           =   8
      Left            =   2685
      TabIndex        =   5
      Top             =   3210
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Co_Cap_Carga_Estatica"
      Height          =   285
      Index           =   9
      Left            =   2685
      TabIndex        =   4
      Top             =   3525
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Pu_Carga_Lim_Fatiga"
      Height          =   285
      Index           =   10
      Left            =   2685
      TabIndex        =   3
      Top             =   3855
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "n_Velocidad"
      Height          =   285
      Index           =   11
      Left            =   2685
      TabIndex        =   2
      Top             =   4170
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Masa"
      Height          =   285
      Index           =   12
      Left            =   2685
      TabIndex        =   1
      Top             =   4485
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Tipo_Orificio"
      Height          =   285
      Index           =   13
      Left            =   2685
      TabIndex        =   0
      Top             =   4815
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "ID_Rodamiento:"
      Height          =   255
      Index           =   0
      Left            =   765
      TabIndex        =   27
      Top             =   630
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cil_Designacion:"
      Height          =   255
      Index           =   1
      Left            =   765
      TabIndex        =   26
      Top             =   945
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Marca_Rodam:"
      Height          =   255
      Index           =   3
      Left            =   765
      TabIndex        =   25
      Top             =   1290
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "d_Diam_Interno:"
      Height          =   255
      Index           =   4
      Left            =   765
      TabIndex        =   24
      Top             =   1605
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "D_Diam_Externo:"
      Height          =   255
      Index           =   5
      Left            =   765
      TabIndex        =   23
      Top             =   1935
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "B_Ancho:"
      Height          =   255
      Index           =   6
      Left            =   765
      TabIndex        =   22
      Top             =   2250
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "a_Ang_Contacto:"
      Height          =   255
      Index           =   7
      Left            =   765
      TabIndex        =   21
      Top             =   2565
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "fa_Factor:"
      Height          =   255
      Index           =   8
      Left            =   765
      TabIndex        =   20
      Top             =   2895
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "C_Cap_Carga_Dinamica:"
      Height          =   255
      Index           =   9
      Left            =   765
      TabIndex        =   19
      Top             =   3210
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Co_Cap_Carga_Estatica:"
      Height          =   255
      Index           =   10
      Left            =   765
      TabIndex        =   18
      Top             =   3525
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Pu_Carga_Lim_Fatiga:"
      Height          =   255
      Index           =   11
      Left            =   765
      TabIndex        =   17
      Top             =   3855
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "n_Velocidad:"
      Height          =   255
      Index           =   12
      Left            =   765
      TabIndex        =   16
      Top             =   4170
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Masa:"
      Height          =   255
      Index           =   13
      Left            =   765
      TabIndex        =   15
      Top             =   4485
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tipo_Orificio:"
      Height          =   255
      Index           =   14
      Left            =   765
      TabIndex        =   14
      Top             =   4815
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

