VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmInsertarEXE 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11115
   Icon            =   "frmInsertarEXE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   7185
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   420
      Left            =   2835
      TabIndex        =   2
      Top             =   4860
      Width           =   2040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   4860
      Width           =   2445
   End
   Begin VB.PictureBox Picture1 
      Height          =   4065
      Left            =   0
      ScaleHeight     =   4005
      ScaleWidth      =   5220
      TabIndex        =   0
      Top             =   0
      Width           =   5280
   End
End
Attribute VB_Name = "frmInsertarEXE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Incrusta
Private Sub Command1_Click()
Call Incrustar_calculadora("calc.exe", Picture1, "Calculadora", Me)
End Sub
  
'Libera y cierra
Private Sub Command2_Click()
Call Liberar_Programa(El_Hwnd_Programa)
End Sub
  
Private Sub Form_Load()
Me.Caption = " Ejemplo del Api SetParent para" & _
             "incrustar la calculadora en un picturebox"
Command1.Caption = " >> Incrustar calculadora "
Command2.Caption = " Liberar y finalizar "
End Sub
  
'Libera y cierra
Private Sub Form_Unload(Cancel As Integer)
Call Cerrar_Programa(El_Hwnd_Programa)
End
End Sub

