VERSION 5.00
Begin VB.Form frmNiv_Sev 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Niveles de Severidad de Vibraciones"
   ClientHeight    =   8265
   ClientLeft      =   7710
   ClientTop       =   1710
   ClientWidth     =   7575
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -60
      TabIndex        =   1
      Top             =   7530
      Width           =   7695
   End
   Begin VB.CommandButton cmdRangos_sev 
      Caption         =   "Rangos de Severidad"
      Height          =   405
      Left            =   270
      TabIndex        =   0
      Top             =   7770
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   7485
      Left            =   0
      Picture         =   "frmNiv_Sev.frx":0000
      Top             =   0
      Width           =   7590
   End
End
Attribute VB_Name = "frmNiv_Sev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRangos_sev_Click()
frmRangos_Sev.Show vbModal
End Sub
