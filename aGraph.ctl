VERSION 5.00
Begin VB.UserControl aGraph 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   6030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895
   ScaleHeight     =   6030
   ScaleWidth      =   8895
   Begin VB.PictureBox PaintBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   840
      ScaleHeight     =   4665
      ScaleWidth      =   7785
      TabIndex        =   0
      Top             =   30
      Width           =   7815
      Begin VB.Label PointBlip 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   1800
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Line GraphLine 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         Index           =   0
         Visible         =   0   'False
         X1              =   600
         X2              =   3840
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line AltLine 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Index           =   1
         X1              =   600
         X2              =   7800
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line AltLine 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Index           =   2
         X1              =   600
         X2              =   7800
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line AltLine 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Index           =   3
         X1              =   600
         X2              =   7800
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line AltLine 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Index           =   4
         X1              =   600
         X2              =   7800
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line AltLine 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Index           =   5
         X1              =   600
         X2              =   7800
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line AltLine 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Index           =   6
         X1              =   600
         X2              =   7800
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line AltLine 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Index           =   7
         X1              =   600
         X2              =   7800
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line AltLine 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Index           =   8
         X1              =   600
         X2              =   7800
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line AltLine 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Index           =   9
         X1              =   600
         X2              =   7800
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line AltLine 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Index           =   0
         X1              =   600
         X2              =   7800
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.Label AltLabel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   50
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label AltLabel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   50
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label AltLabel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   50
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label AltLabel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   50
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label AltLabel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   50
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label AltLabel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   50
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label AltLabel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   50
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label AltLabel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   50
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label AltLabel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   50
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label AltLabel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   50
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "aGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'misma vieja historia, no pudo encontrar la altitud / gráfico de control, así que cortó un conjunto
'Estoy trabajando en una utilidad de GPS y necesitaba una forma gráfica de mostrar la altitud de los puntos
'con un poco de información básica ..

'cosas que me gustaría añadir, pero no tiene tiempo para el momento:
'asignado por el usuario de fondo, las líneas de división, y los puntos
'asignado por el usuario unidad de medida (m, m, m, m, etc)
'si es posible realizar todo el control un poco más rápido, parece que tiene un cierto retraso.
'añadir" globo "sobre herramientas, con información más completa
'imagen asignable para el fondo, no sólo de color sólido
'comprobación de errores ... hacer que todos podemos hacer un mejor trabajo escrito de código libre de errores!

'si usted hace las mejoras por favor subirlo, o por correo electrónico me salsa117@hotmail.com.
'No estoy buscando ningún crédito, lo que este código es tuyo que ver con lo que usted desea.

Public Values As New Collection
Public aDescription As New Collection
Public MaxValue As Integer
Dim GraphLineColors As ColorConstants

Private Sub UserControl_Initialize()
    If MaxValue < 10 Then
        MaxValue = 10
    End If
End Sub

Private Sub UserControl_Resize()
PaintBox.Left = 700
PaintBox.Top = 100
PaintBox.Width = (UserControl.Width - 750)
PaintBox.Height = (UserControl.Height - 100)

    For i = 0 To 9
    
        'setup horizontal lines for Grid
        'first set their left and width
        AltLine(i).X1 = 0
        AltLine(i).X2 = (UserControl.Width)
        'then set their spacing
        AltLine(i).Y1 = ((UserControl.Height / 10) * i) - 50
        AltLine(i).Y2 = ((UserControl.Height / 10) * i) - 50
        
        'show the label on the left hand side
        AltLabel(i).Top = ((UserControl.Height / 10) * i)
        AltLabel(i).Visible = True
        AltLabel(i).BackStyle = 0
        
        'tiene que haber una manera mejor de hacer esto...
        Select Case True
            Case i = 0
                AltLabel(i).Caption = (MaxValue / 10) * 10 & ".0" 'make the "ft" user assignable, M, m, ft, yd, etc
            Case i = 1
                AltLabel(i).Caption = (MaxValue / 10) * 9 & ".0"
            Case i = 2
                AltLabel(i).Caption = (MaxValue / 10) * 8 & ".0"
            Case i = 3
                AltLabel(i).Caption = (MaxValue / 10) * 7 & ".0"
            Case i = 4
                AltLabel(i).Caption = (MaxValue / 10) * 6 & ".0"
            Case i = 5
                AltLabel(i).Caption = (MaxValue / 10) * 5 & ".0"
            Case i = 6
                AltLabel(i).Caption = (MaxValue / 10) * 4 & ".0"
            Case i = 7
                AltLabel(i).Caption = (MaxValue / 10) * 3 & ".0"
            Case i = 8
                AltLabel(i).Caption = (MaxValue / 10) * 2 & ".0"
            Case i = 9
                AltLabel(i).Caption = (MaxValue / 10) * 1 & ".0"
        End Select
            
    Next i

End Sub

Sub ReDraw()
On Error Resume Next
    Hspace = (PaintBox.Width / Val(Values.Count))
    Vspace = (PaintBox.Height / MaxValue)
    YStart = (PaintBox.Height)
    For i = 0 To Values.Count
        On Error Resume Next
        If i = 0 Then
        'start the line at 0, in the middle
            LastX = -1
            LastY = (PaintBox.Height)
        End If
            'load the current line
            Load GraphLine(i)
            'set the color
            GraphLine(i).BorderColor = GraphLineColors
            'x1 should match last line's end point, x2 is new
            GraphLine(i).X1 = LastX - 1
            GraphLine(i).X2 = (Hspace * i)
            'y1 should match last line's end point, y2 is new
            GraphLine(i).Y1 = LastY
            GraphLine(i).Y2 = (YStart - Values.Item(i) * Vspace)
            'set the values for the next line
            LastX = (Hspace * i + 1)
            LastY = (YStart - Values.Item(i) * Vspace)
            GraphLine(i).Visible = True
            'start showing points
            Load PointBlip(i)
            PointBlip(i).Left = (LastX - 50)
            PointBlip(i).Top = (LastY)
            PointBlip(i).ToolTipText = i & " = " & aDescription.Item(i)
            'make sure the the blips show on top of the lines
            PointBlip(i).ZOrder
            PointBlip(i).Visible = True
    Next i
End Sub

'------------------------------------------------------------------------------------------------
Public Property Get GraphLineColor() As OLE_COLOR
    GraphLineColor = GraphLineColors
End Property

Public Property Let GraphLineColor(newGraphLineColor As OLE_COLOR)
    GraphLineColors = newGraphLineColor
    PropertyChanged "GraphLineColor"
End Property


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        GraphLineColors = PropBag.ReadProperty("GraphLineColor", vbRed)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("GraphLineColor", GraphLineColors, vbRed)
End Sub
