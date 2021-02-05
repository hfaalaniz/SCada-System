VERSION 5.00
Begin VB.Form frmReporteRegistros 
   Caption         =   "Reporte"
   ClientHeight    =   9525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   ScaleHeight     =   9525
   ScaleWidth      =   13110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   405
      Left            =   10350
      TabIndex        =   4
      Top             =   6270
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   6270
      Width           =   9495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   345
      Left            =   600
      TabIndex        =   2
      Top             =   5640
      Width           =   8745
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2655
      Left            =   12540
      TabIndex        =   1
      Top             =   1110
      Width           =   285
   End
   Begin VB.PictureBox picR 
      Height          =   4065
      Left            =   0
      ScaleHeight     =   4005
      ScaleWidth      =   8175
      TabIndex        =   0
      Top             =   0
      Width           =   8235
   End
End
Attribute VB_Name = "frmReporteRegistros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CYHSCROLL = 3
Private Const SM_CXVSCROLL = 2
'variable para el recordset y la coenxión Ado
Dim cn As ADODB.connection
Dim rst As ADODB.Recordset
' variable de tipo Pictuerbox que es para el reporte
Private picReporte As PictureBox
' SubRutina que imprime los regisros en el picturebox _
  para crear el reporte ( Recibe el recordset y el picturebox )
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowReport(rs As ADODB.Recordset, _
                       Contenedor As Object, _
                       Optional TextHeader As String, _
                       Optional TextHeader2 As String, _
                       Optional TextFooter As String, _
                       Optional FontColorHeader As Long, _
                       Optional FontColorHeader2 As Long, _
                       Optional FontColorFooter As Long, _
                       Optional ForeColorRegistros As Long, _
                       Optional ForeColorField As Long, _
                       Optional Margen As Single = 100)
   
On Error GoTo Error_Sub
   
   
   Dim Posy As Long
   Dim Posy2 As Long
   Dim posx As Long
   Dim AltoLetra As Long
   Dim Columna As Integer
   Dim Dato As String
   Dim MaxPosx As Long
   Dim Campo As Boolean
   Dim AltoPicture As Double
   Dim t As Long
   
   Me.MousePointer = vbHourglass
   ' crea un picturebox
   If picReporte Is Nothing Then
      Set picReporte = Me.controls.Add("vb.PictureBox", "Pic1")
   End If
   ' borra el contenido del picture
   picReporte.Cls
   ' para que se actualice mas rápido
   LockWindowUpdate picReporte.hWnd
   ' propiedades varias
   With picReporte
        Set .Container = Contenedor
        .Width = 3000 * Screen.TwipsPerPixelX
        .Appearance = 0
        .BorderStyle = 1
        .Visible = True
        .BackColor = vbWhite
        .AutoRedraw = True
        .CurrentX = Margen
        .CurrentY = 210
        .FontName = "Arial"
        .FontSize = 12
        .ForeColor = FontColorHeader
   End With
   AltoLetra = picReporte.TextHeight("A")
   picReporte.Height = (AltoLetra * 2) + picReporte.CurrentY
   ' Enzabezado
   picReporte.Print TextHeader
   ' Encabezado 2
   With picReporte
        .FontSize = 8
        .Height = (AltoLetra * 2) + .CurrentY
        .ForeColor = FontColorHeader2
        .CurrentX = Margen
        picReporte.Print TextHeader2
        picReporte.Line (Margen, .CurrentY)-(3000 * Screen.TwipsPerPixelX, .CurrentY)
   End With
   
   Posy2 = picReporte.CurrentY
   ' recorre las columnas
   For Columna = 0 To rs.Fields.Count - 1
   Campo = True
   rst.MoveFirst
      
   Posy = Posy2 + AltoLetra + Margen ' currenty para escribir
   posx = (posx + MaxPosx) + 500
   MaxPosx = 0
    
   picReporte.Height = ((rst.RecordCount) * (AltoLetra)) + (Posy * 2)
   'Recorre los registros
   Do While Not rs.EOF
        
        If IsNull(rst(Columna).value) Then
            Dato = vbNullString
        Else
            Dato = rst(Columna).value
        End If

        If picReporte.TextWidth(Dato) > MaxPosx Then
           MaxPosx = picReporte.TextWidth(Dato)
        End If
        
        If Campo Then
        
           If picReporte.TextWidth(rst(Columna).Name) > MaxPosx Then
              MaxPosx = picReporte.TextWidth(rst(Columna).Name)
           End If
           ' propiedades de fuente para los campos
           With picReporte
              .FontBold = True
              .FontUnderline = True
              .ForeColor = ForeColorField
              .FontSize = 8
           End With
           ' imprime el texto de los campos
           Imprimir_dato posx, Posy, rst(Columna).Name
           Posy = Posy + AltoLetra
           Campo = False
           ' propiedades de fuente para los registros
           With picReporte
              .FontBold = False
              .ForeColor = ForeColorRegistros
              .FontSize = 8
              .FontUnderline = False
           End With
        Else
           ' imprime el dato del registro actual
           Call Imprimir_dato(posx, Posy, Dato)
        
           rs.MoveNext ' siguiente registro
           ' Establece el CurrentY
           Posy = Posy + AltoLetra
        End If
   Loop
   t = t + MaxPosx + 250 + (Margen * 4)
   Next
   If t < Printer.ScaleWidth Then
      picReporte.Width = Printer.ScaleWidth
   Else
      picReporte.Width = t
   End If
   ' Fin de los registros
   '''''''''''''''''''''''''''''
   ' imprime el Pie de página
   picReporte.Print ""
   picReporte.Print ""
   
   picReporte.ForeColor = FontColorFooter
   picReporte.Line (Margen, picReporte.CurrentY)-(3000 * Screen.TwipsPerPixelX, picReporte.CurrentY)
   picReporte.CurrentX = Margen
   picReporte.CurrentY = Posy + AltoLetra
   picReporte.Print TextFooter
   
   LockWindowUpdate False
   picReporte.Refresh
   ' Establece las barras de scroll
   ShowScrollBar picR, picReporte
   Me.MousePointer = 0
Exit Sub
' error
Error_Sub:
Me.MousePointer = 0
LockWindowUpdate False
MsgBox Err.Description, vbCritical
End Sub

Private Sub Imprimir_dato(x As Long, y As Long, s As String)
   picReporte.CurrentX = x + 250
   picReporte.CurrentY = y
   picReporte.Print s;
End Sub

Private Sub Command1_Click()
    ' carga el recordset
    rst.Open Text1.Text, cn, adOpenStatic
    ShowReport rst, picR, _
               "Ejemplo simple de reporte", _
               "Total : (" & rst.RecordCount & " registros)", _
               "Fin del listado - Fecha " & Date, _
               &H8000000D, &H808080, &H808080, &H8000000D, 8421504
    rst.Close
End Sub

Private Sub ShowScrollBar(picContenedor As Object, PicData As PictureBox)
    Dim W As Single
    Dim H As Single
    
    Set VScroll1.Container = picContenedor
    Set HScroll1.Container = picContenedor
    ' Izquierda y top del picture que tiene el reporte
    PicData.Move 70, 80
    ' redimensiona los scrollbar
    W = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    H = GetSystemMetrics(SM_CYHSCROLL) * Screen.TwipsPerPixelY
    HScroll1.Move 0, picContenedor.ScaleHeight - H, picContenedor.ScaleWidth - W, H
    
    VScroll1.Move picContenedor.ScaleWidth - W, 0, W, picContenedor.ScaleHeight
    ' Asigna los valores Max y Largchange de las barras de desplazamiento
    HScroll1.LargeChange = 15
    VScroll1.LargeChange = 15
    
    HScroll1.Max = (PicData.Width - picContenedor.ScaleWidth + VScroll1.Width) / 120 + 1
    VScroll1.Max = (PicData.Height - picContenedor.ScaleHeight + HScroll1.Height) / 120 + 1

    If PicData.Height <= picContenedor.Height Then
       VScroll1.Visible = False
       HScroll1.Width = picContenedor.ScaleWidth
    Else
       VScroll1.Visible = True
    End If
    
    If PicData.Width <= picContenedor.Width Then
       HScroll1.Visible = False
       VScroll1.Height = picContenedor.ScaleHeight
    Else
       HScroll1.Visible = True
    End If
    HScroll1.value = 0
    VScroll1.value = 0
    ' trae los controles hacia el frente
    HScroll1.ZOrder 0
    VScroll1.ZOrder 0
End Sub

' eventos de los ScroolBar
'''''''''''''''''''''''''''''''''
Private Sub HScroll1_Change()
    picReporte.Left = (-HScroll1.value * 120) + 70
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub VScroll1_Change()
    picReporte.Top = (-CSng(VScroll1.value) * 120) + 80
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub

'''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    ' nueva conexión Ado
    Set cn = New ADODB.connection
    cn.CursorLocation = adUseClient
    ' cadena de conexión (us la base de datos biblio del directorio de visual basic )
    'cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
    '                      App.Path & "\vibramec.MDB;Persist Security Info=False"
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BIBLIO2003.MDB;Persist Security Info=False"
    ' abre la conexión
    cn.Open
    ' Crear un nuevo recordset
    Set rst = New ADODB.Recordset
    Command1.Caption = "Mostrar reporte"
    Text1.Text = "Select * From Publishers"
    VScroll1.Visible = False
    HScroll1.Visible = False
End Sub



