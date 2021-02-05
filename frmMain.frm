VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Espectro de frecuencias de la entrada"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "Parar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3135
      TabIndex        =   11
      Top             =   5865
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Iniciar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1860
      TabIndex        =   10
      Top             =   5865
      Width           =   1215
   End
   Begin VB.PictureBox picSpectrum 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3735
      Left            =   60
      ScaleHeight     =   247
      ScaleMode       =   0  'User
      ScaleWidth      =   1024
      TabIndex        =   9
      Top             =   2025
      Width           =   13185
   End
   Begin VB.Frame frameRecSettings 
      Caption         =   "Configuración de Entradas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   4365
      Begin VB.PictureBox picRecSettings 
         BorderStyle     =   0  'None
         Height          =   1440
         Left            =   75
         ScaleHeight     =   1440
         ScaleWidth      =   4215
         TabIndex        =   1
         Top             =   225
         Width           =   4215
         Begin VB.ComboBox cboDevice 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1125
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   105
            Width           =   2865
         End
         Begin VB.ComboBox cboMixerLine 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1125
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   525
            Width           =   2265
         End
         Begin VB.CommandButton cmdAutoSelectMic 
            Caption         =   "AMA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3450
            TabIndex        =   2
            ToolTipText     =   "Automatische Mikrophon Auswahl"
            Top             =   525
            Width           =   540
         End
         Begin MSComctlLib.Slider sldVolume 
            Height          =   390
            Left            =   1125
            TabIndex        =   3
            Top             =   975
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   688
            _Version        =   393216
            LargeChange     =   2000
            SmallChange     =   500
            Max             =   65535
            SelStart        =   65535
            TickStyle       =   3
            Value           =   65535
         End
         Begin VB.Label lblDevice 
            AutoSize        =   -1  'True
            Caption         =   "Dispositivo:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   8
            Top             =   150
            Width           =   825
         End
         Begin VB.Label lblMixerLine 
            AutoSize        =   -1  'True
            Caption         =   "Canal:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   165
            TabIndex        =   7
            Top             =   570
            Width           =   465
         End
         Begin VB.Label lblMixerLineVolume 
            AutoSize        =   -1  'True
            Caption         =   "Volumen:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   6
            Top             =   975
            Width           =   660
         End
      End
   End
   Begin VB.Label lblFreqHz 
      AutoSize        =   -1  'True
      Caption         =   "0 Hz"
      Height          =   195
      Left            =   945
      TabIndex        =   13
      Top             =   5865
      Width           =   330
   End
   Begin VB.Label lblFreq 
      AutoSize        =   -1  'True
      Caption         =   "Frequencia:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   5865
      Width           =   840
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Samplerate            As Long = 44100 '22050
Private Const Channels              As Long = 2         ' No hay tratamiento para 2 canales construido en!
Private Const SAMPLEWIDTH           As Long = 2         ' 16-bits muestras no se puede cambiar
Private Const BLOCKALIGN            As Long = Channels * SAMPLEWIDTH
Private Const BYTESPERSEC           As Long = Samplerate * Channels * 2

Private Const BUFFERLENGTH          As Long = 25        ' Frecuencia de actualización en ms
Private Const FFTSIZE               As Long = 512       ' Elementos, los poderes únicos de dos posibles
Private Const FFTMAXMAGNITUDE       As Double = 0.09  '0.2     ' "Zoom", menor es el mejor
'Private Const PI                    As Double = 3.14159265358979

Private WithEvents m_clsRecorder    As WaveInRecorder
Attribute m_clsRecorder.VB_VarHelpID = -1
Private m_blnRecording              As Boolean

Private m_fftwPlan                  As FFTWPlan

Private m_sngSpectrum(FFTSIZE - 1)  As Single
Private m_sngWindow(FFTSIZE - 1)    As Single

Private m_sngFFTIn(FFTSIZE - 1)     As Single
Private m_sngFFTOut(FFTSIZE - 0)    As Single

Private Sub m_clsRecorder_GotData(intBuffer() As Integer, lngLen As Long)
    Dim dblValue                    As Double
    Dim i                           As Long
    Dim lngHeight                   As Long
    Dim lngWidth                    As Long
    
    With picSpectrum
        lngHeight = .ScaleHeight
        lngWidth = .ScaleWidth
    End With
    
    ' Hanning ventana para aplicar a las muestras para reducir el ruido
    For i = 0 To FFTSIZE - 1
        m_sngFFTIn(i) = intBuffer(i) * m_sngWindow(i)
    Next
    
    ' Las muestras de transformada de tiempo-frecuencia en dominio
    FFTW_Execute m_fftwPlan
    
    picSpectrum.Cls
    
    For i = 0 To FFTSIZE - 1 Step 2
        ' Magnitud de la banda de frecuencia
        dblValue = Sqr(m_sngFFTOut(i + 0) * m_sngFFTOut(i + 0) + m_sngFFTOut(i + 1) * m_sngFFTOut(i + 1))

        dblValue = dblValue / (FFTSIZE / 4) / 32768
        
        ' aumentar el volumen de
        If dblValue > FFTMAXMAGNITUDE Then
            dblValue = FFTMAXMAGNITUDE
        Else
            dblValue = dblValue / FFTMAXMAGNITUDE
        End If
        
        ' Promedio ponderado de un movimiento más suave del espectro
        m_sngSpectrum(i / 2) = 0.32 * m_sngSpectrum(i / 2) + 0.68 * dblValue
        
        picSpectrum.Line (i / 2, lngHeight)-(i / 2, lngHeight - m_sngSpectrum(i / 2) * lngHeight)
    Next
End Sub

Private Sub ShowMixerLines()
    Dim i   As Long
    
    cboMixerLine.Clear
    
    For i = 0 To m_clsRecorder.MixerLineCount - 1
        cboMixerLine.AddItem m_clsRecorder.MixerLineName(i)
    Next
    
    If m_clsRecorder.MixerLineCount > 0 Then
        cboMixerLine.ListIndex = 0
    Else
        'MsgBox "No hay ningún dispositivo de grabación de dos canales de salida!", vbExclamation
    End If
End Sub

Private Sub VerDispositivos()
    Dim i   As Long
    
    cboDevice.Clear
    
    For i = 0 To m_clsRecorder.DeviceCount - 1
        cboDevice.AddItem m_clsRecorder.DeviceName(i)
    Next
    
    If m_clsRecorder.DeviceCount > 0 Then
        cboDevice.ListIndex = 0
    Else
        MsgBox "No se econtraron dispositivos que puedan grabar!", vbExclamation
    End If
End Sub

Private Property Get Recording() As Boolean
    Recording = m_blnRecording
End Property

Private Property Let Recording(ByVal blnValue As Boolean)
    m_blnRecording = blnValue
    
    cmdStart.Enabled = Not blnValue
    cmdStop.Enabled = blnValue
    
    frameRecSettings.Enabled = Not blnValue
End Property

Private Sub cboDevice_Click()
    cboMixerLine.Clear
    
    If Not m_clsRecorder.SelectDevice(cboDevice.ListIndex) Then
        MsgBox "No se puede seleccionar el dispositivo!", vbExclamation
    Else
        ShowMixerLines
    End If
End Sub

Private Sub cboMixerLine_Click()
    If Not m_clsRecorder.SelectMixerLine(cboMixerLine.ListIndex) Then
        MsgBox "No se pudo seleccionar el canal a grabar!", vbExclamation
    Else
        sldVolume.value = m_clsRecorder.MixerLineVolume
    End If
End Sub

Private Sub cmdAutoSelectMic_Click()
    Dim i   As Long

    For i = 0 To m_clsRecorder.MixerLineCount - 1
        If m_clsRecorder.MixerLineType(i) = MIXERLINE_MICROPHONE Then
            cboMixerLine.ListIndex = i
            Exit For
        End If
    Next
    
    If i = m_clsRecorder.MixerLineCount Then
        MsgBox "No encontró el micrófono!", vbExclamation
    End If
End Sub

Private Sub cmdStart_Click()
    With m_clsRecorder
        .BufferSize = AlignToBlockAlign(BUFFERLENGTH / 1000 * BYTESPERSEC)
        If Not .StartRecord(Samplerate, Channels) Then
            MsgBox "No se pudo iniciar la grabación!", vbExclamation
        Else
            Recording = True
        End If
    End With
End Sub

Private Sub cmdStop_Click()
    If Not m_clsRecorder.StopRecord() Then
        MsgBox "No se ha podido detener la grabación!", vbExclamation
    Else
        Recording = False
    End If
End Sub

Private Function AlignToBlockAlign(ByVal Bytes As Long) As Long
    AlignToBlockAlign = Bytes - (Bytes Mod BLOCKALIGN)
End Function

Private Sub picSpectrum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < FFTSIZE \ 2 - 1 Then
        lblFreqHz.Caption = " " & CLng(Samplerate / FFTSIZE * X) & " Hz"
    End If
End Sub

Private Sub sldVolume_Change()
    m_clsRecorder.MixerLineVolume = sldVolume.value
End Sub

Private Sub sldVolume_Scroll()
    m_clsRecorder.MixerLineVolume = sldVolume.value
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    Set m_clsRecorder = New WaveInRecorder
    
    Recording = False
    VerDispositivos
    
    For i = 0 To FFTSIZE - 1
        m_sngWindow(i) = HanningWindow(i, FFTSIZE)
    Next
    
    FFTWInit
    m_fftwPlan = FFTW_Create_Plan_r2c_1d(FFTSIZE, VarPtr(m_sngFFTIn(0)), VarPtr(m_sngFFTOut(0)), FFTW_MEASURE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_clsRecorder.StopRecord
    Recording = False
    
    FFTW_Destroy_Plan m_fftwPlan
    FFTWTerm
End Sub
