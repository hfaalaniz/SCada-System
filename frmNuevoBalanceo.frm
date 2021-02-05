VERSION 5.00
Begin VB.Form frmNuevoBalanceo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NuevoBalanceo"
   ClientHeight    =   4500
   ClientLeft      =   6645
   ClientTop       =   1995
   ClientWidth     =   6810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFirst 
      Height          =   300
      Left            =   120
      Picture         =   "frmNuevoBalanceo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdPrevious 
      Height          =   300
      Left            =   465
      Picture         =   "frmNuevoBalanceo.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdNext 
      Height          =   300
      Left            =   1980
      Picture         =   "frmNuevoBalanceo.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdLast 
      Height          =   300
      Left            =   2325
      Picture         =   "frmNuevoBalanceo.frx":09C6
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   -60
      TabIndex        =   17
      Top             =   3900
      Width           =   7155
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   5640
      TabIndex        =   16
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   300
      Left            =   5640
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Modificar"
      Height          =   300
      Left            =   4080
      TabIndex        =   14
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Guardar"
      Height          =   300
      Left            =   2940
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Nuevo"
      Height          =   300
      Left            =   2940
      TabIndex        =   12
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton btnNuevaE 
      Caption         =   ">>"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4920
      TabIndex        =   9
      Top             =   2220
      Width           =   465
   End
   Begin VB.ComboBox cmbEmpresa 
      Appearance      =   0  'Flat
      DataField       =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2010
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   2220
      Width           =   2835
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      DataField       =   "Equipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2010
      TabIndex        =   3
      Top             =   3210
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Empresa"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   7
      Top             =   2265
      Width           =   2445
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      DataField       =   "Ubicacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2010
      TabIndex        =   1
      Top             =   1425
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      DataField       =   "ID_NuevoBal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2190
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   810
      TabIndex        =   22
      Top             =   4080
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   420
      TabIndex        =   11
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese los datos de configuración del balanceo."
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
      Left            =   1350
      TabIndex        =   10
      Top             =   570
      Width           =   4350
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5940
      Picture         =   "frmNuevoBalanceo.frx":0D08
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Equipo:"
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
      Index           =   3
      Left            =   1215
      TabIndex        =   8
      Top             =   3210
      Width           =   690
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Empresa:"
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
      Index           =   2
      Left            =   1035
      TabIndex        =   6
      Top             =   2265
      Width           =   870
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Ubicacion:"
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
      Index           =   1
      Left            =   945
      TabIndex        =   5
      Top             =   1425
      Width           =   960
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ID_NuevoBal:"
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
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000A&
      BorderColor     =   &H8000000A&
      FillColor       =   &H8000000A&
      FillStyle       =   0  'Solid
      Height          =   915
      Left            =   0
      Top             =   0
      Width           =   10395
   End
End
Attribute VB_Name = "frmNuevoBalanceo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db As connection
Dim adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim adoEmp As Recordset
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim cantReg As Integer

Private Sub btnNuevaE_Click()
adoEmp.Close 'cierra el recordset para actualizarlo
frmEmpresas.Show vbModal

Set adoEmp = New Recordset
adoEmp.Open "SELECT Empresa FROM Empresas ORDER BY Empresa", db, adOpenStatic, adLockOptimistic
adoEmp.Requery
adoEmp.MoveFirst
Call Cargar(cmbEmpresa, adoEmp, "Empresa")
If adoPrimaryRS!Empresa = Not Null Then
    cmbEmpresa.Text = adoPrimaryRS!Empresa
Else
    cmbEmpresa.Text = "    "
End If
cmbEmpresa.Text = txtFields(2).Text
End Sub

Private Sub Form_Load()
  
  Set db = New connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VibraMec\VIBRAMEC.mdb;"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "SELECT ID_NuevoBal,Ubicacion,Empresa,Equipo FROM NuevoBalanceo Order by Ubicacion", db, adOpenStatic, adLockOptimistic
  ' lo abre y le pasa el sql
  Set adoEmp = New Recordset
  adoEmp.Open "SELECT Empresa FROM Empresas ORDER BY Empresa", db, adOpenStatic, adLockOptimistic
  Call Cargar(cmbEmpresa, adoEmp, "Empresa")
  cantReg = adoPrimaryRS.RecordCount
  
  Dim oText As TextBox
  'Enlaza los cuadros de texto con el proveedor de datos
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  
  adoPrimaryRS.MoveFirst
  If adoPrimaryRS!Empresa = Not Null Then
    cmbEmpresa.Text = adoPrimaryRS!Empresa
  Else
    cmbEmpresa.Text = " "
  End If
    If adoPrimaryRS!Empresa = Null Then
        cmbEmpresa.Text = " "
    Else
        cmbEmpresa.Text = adoPrimaryRS!Empresa
    End If
  mbDataChanged = False

Set Tooltip = New cToolTip
Tooltip.Create btnNuevaE, "Haga click aqui para agregar una nueva empresa.", _
                             TTBalloonIfActive, False, TTIconInfo, "Agregar", vbBlack, vbWhite, 100, 20000
Tooltips.Add Tooltip, btnNuevaE.Name 'no te olvides de mantenerlo

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
    cmbEmpresa.Text = ""
    lblStatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
  End With
  txtFields(0).Text = cantReg + 1
  Exit Sub
AddErr:
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
  cmbEmpresa.Text = txtFields(2).Text
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'va al nuevo registro
  End If
  
  adoPrimaryRS!Empresa = cmbEmpresa.Text
  
  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  cantReg = adoPrimaryRS.RecordCount
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
  
  'muestra el registro actual
    If adoPrimaryRS!Empresa = Null Then
        cmbEmpresa.Text = "    "
    Else
        cmbEmpresa.Text = adoPrimaryRS!Empresa
    End If
  
  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False
  
  'muestra el registro actual
    If adoPrimaryRS!Empresa = Null Then
        cmbEmpresa.Text = "    "
    Else
        cmbEmpresa.Text = adoPrimaryRS!Empresa
    End If
  
  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error Resume Next  'GoNextError

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'ha sobrepasado el final; vuelva atrás
    adoPrimaryRS.MoveLast
  End If
  'muestra el registro actual
    If adoPrimaryRS!Empresa = Null Then
        cmbEmpresa.Text = "    "
    Else
        cmbEmpresa.Text = adoPrimaryRS!Empresa
    End If
    
  mbDataChanged = False

'  Exit Sub
'GoNextError:
'  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error Resume Next  ' GoPrevError

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'ha sobrepasado el final; vuelva atrás
    adoPrimaryRS.MoveFirst
  End If
  'muestra el registro actual
    If adoPrimaryRS!Empresa = Null Then
        cmbEmpresa.Text = "    "
    Else
        cmbEmpresa.Text = adoPrimaryRS!Empresa
    End If
  'muestra el registro actual
  mbDataChanged = False

'  Exit Sub
'GoPrevError:
'  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  btnNuevaE.Enabled = Not bVal
  cmdClose.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub


Private Sub cmbEmpresa_Change()
    'Le pasamos el ComboBox que queremos, en este caso un cmbEmpresa
    Autocompletar_Combo cmbEmpresa
End Sub
  
Private Sub cmbEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
  
    Select Case KeyCode
        'Si la tecla presionada es Backspace o la tecla Delete
        Case vbKeyBack, vbKeyDelete
            Select Case Len(cmbEmpresa.Text)
                Case Is <> 0
                    KeyRetroceso = True
  
            End Select
    End Select
End Sub
  
'A este procedimento le enviamos como _
parámetro el Control Combo que queremos utilizar.
Public Function Autocompletar_Combo(Combo As ComboBox)
  
Dim i As Integer, posSelect As Integer
  
    Select Case (KeyRetroceso Or Len(Combo.Text) = 0)
        Case True
            KeyRetroceso = False
            Exit Function
    End Select
  
    With Combo
  
    'Recorremos todos los elementos del combo
    For i = 0 To .ListCount - 1
        'Si hay coincidencia
        If InStr(1, .List(i), .Text, vbTextCompare) = 1 Then
            posSelect = .SelStart
            'Mostramos el texto en el combo
            .Text = .List(i)
            'Indicamos el comienzo de la selección
            .SelStart = posSelect
            'Acá seleccionamos el texto
            .SelLength = Len(.Text) - posSelect
  
            Exit For
        End If
    Next i
  
    End With
End Function
  
'Este procedimiento es para ocultar o desplegar _
el combo cuando presionamos el enter
Private Sub cmbEmpresa_KeyPress(KeyAscii As Integer)
  
Dim resp As Integer
  
    If KeyAscii = 13 Then
        'Si le pasamos a SendMessageLong el valor False lo cierra
        resp = SendMessageLong(cmbEmpresa.hWnd, &H14F, False, 0)
    Else
        'si le pasamos True a SendMessageLong lo adespliega, es decir cuando
        'presionamos una tecla diferente al Enter
        resp = SendMessageLong(cmbEmpresa.hWnd, &H14F, True, 0)
    End If
End Sub

