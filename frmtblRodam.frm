VERSION 5.00
Begin VB.Form frmtblRodam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "tblRodam"
   ClientHeight    =   7200
   ClientLeft      =   1095
   ClientTop       =   375
   ClientWidth     =   16650
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   16650
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   16650
      TabIndex        =   34
      Top             =   6600
      Width           =   16650
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   300
         Left            =   1213
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "A&ctualizar"
         Height          =   300
         Left            =   59
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cerrar"
         Height          =   300
         Left            =   4675
         TabIndex        =   39
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Reno&var"
         Height          =   300
         Left            =   3521
         TabIndex        =   38
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   2367
         TabIndex        =   37
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edición"
         Height          =   300
         Left            =   1213
         TabIndex        =   36
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Ag&regar"
         Height          =   300
         Left            =   59
         TabIndex        =   35
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
      ScaleWidth      =   16650
      TabIndex        =   28
      Top             =   6900
      Width           =   16650
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frmtblRodam.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmtblRodam.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmtblRodam.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmtblRodam.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   33
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "AD1"
      Height          =   285
      Index           =   13
      Left            =   2355
      TabIndex        =   27
      Top             =   5745
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Masa"
      Height          =   285
      Index           =   12
      Left            =   2355
      TabIndex        =   25
      Top             =   5430
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "RPMlim"
      Height          =   285
      Index           =   11
      Left            =   2355
      TabIndex        =   23
      Top             =   5115
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "RPMref"
      Height          =   285
      Index           =   10
      Left            =   2355
      TabIndex        =   21
      Top             =   4785
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Pu"
      Height          =   285
      Index           =   9
      Left            =   2355
      TabIndex        =   19
      Top             =   4470
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Co"
      Height          =   285
      Index           =   8
      Left            =   2355
      TabIndex        =   17
      Top             =   4155
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1"
      Height          =   285
      Index           =   7
      Left            =   2355
      TabIndex        =   15
      Top             =   3825
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C"
      Height          =   285
      Index           =   6
      Left            =   2355
      TabIndex        =   13
      Top             =   3510
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "B"
      Height          =   285
      Index           =   5
      Left            =   2355
      TabIndex        =   11
      Top             =   3195
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "D_Exterior"
      Height          =   285
      Index           =   4
      Left            =   2355
      TabIndex        =   9
      Top             =   2865
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "d_Interior"
      Height          =   285
      Index           =   3
      Left            =   2355
      TabIndex        =   7
      Top             =   2550
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Nombre"
      Height          =   285
      Index           =   2
      Left            =   2355
      TabIndex        =   5
      Top             =   2235
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Marca"
      Height          =   285
      Index           =   1
      Left            =   2355
      TabIndex        =   3
      Top             =   1905
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "IDRodam"
      Height          =   285
      Index           =   0
      Left            =   2355
      TabIndex        =   1
      Top             =   1590
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "AD1:"
      Height          =   255
      Index           =   13
      Left            =   435
      TabIndex        =   26
      Top             =   5745
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Masa:"
      Height          =   255
      Index           =   12
      Left            =   435
      TabIndex        =   24
      Top             =   5430
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "RPMlim:"
      Height          =   255
      Index           =   11
      Left            =   435
      TabIndex        =   22
      Top             =   5115
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "RPMref:"
      Height          =   255
      Index           =   10
      Left            =   435
      TabIndex        =   20
      Top             =   4785
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Pu:"
      Height          =   255
      Index           =   9
      Left            =   435
      TabIndex        =   18
      Top             =   4470
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Co:"
      Height          =   255
      Index           =   8
      Left            =   435
      TabIndex        =   16
      Top             =   4155
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "C1:"
      Height          =   255
      Index           =   7
      Left            =   435
      TabIndex        =   14
      Top             =   3825
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "C:"
      Height          =   255
      Index           =   6
      Left            =   435
      TabIndex        =   12
      Top             =   3510
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "B:"
      Height          =   255
      Index           =   5
      Left            =   435
      TabIndex        =   10
      Top             =   3195
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "D_Exterior:"
      Height          =   255
      Index           =   4
      Left            =   435
      TabIndex        =   8
      Top             =   2865
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "d_Interior:"
      Height          =   255
      Index           =   3
      Left            =   435
      TabIndex        =   6
      Top             =   2550
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nombre:"
      Height          =   255
      Index           =   2
      Left            =   435
      TabIndex        =   4
      Top             =   2235
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Marca:"
      Height          =   255
      Index           =   1
      Left            =   435
      TabIndex        =   2
      Top             =   1905
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "IDRodam:"
      Height          =   255
      Index           =   0
      Left            =   435
      TabIndex        =   0
      Top             =   1590
      Width           =   1815
   End
End
Attribute VB_Name = "frmtblRodam"
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
  Dim db As connection
  Set db = New connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VibraMec\VIBRAMEC.mdb;"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select IDRodam,Marca,Nombre,d_Interior,D_Exterior,B,C,C1,Co,Pu,RPMref,RPMlim,Masa,AD1 from tblRodam Order by Nombre", db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Enlaza los cuadros de texto con el proveedor de datos
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
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
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
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

