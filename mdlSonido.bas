Attribute VB_Name = "mdlSonido"
Option Explicit
Public Const Muestras As Long = 1024   ' potencia de 2 por FFT!
Public DevHandle As Long
Public InData(0 To 1023) As Byte    '511) As Byte
Public OutData(0 To 1023) As Byte
Public Inited As Boolean
Public MinHeight As Long, MinWidth As Long


Public Type WAVEFORMATEX
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BLOCKALIGN As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Public Type WAVEHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long 'wavehdr_tag
    reserved As Long
End Type

Public Type WAVEINCAPS
    ManufacturerID As Integer      'wMid
    ProductID As Integer       'wPid
    DriverVersion As Long       'MMVERSIONS vDriverVersion
    ProductName(1 To 32) As Byte 'szPname[MAXPNAMELEN]
    Formats As Long
    Channels As Integer
    reserved As Integer
End Type

Public Const WAVE_INVALIDFORMAT = &H0&                 '/* invalid format */
Public Const WAVE_FORMAT_1M08 = &H1&                   '/* 11.025 kHz, Mono,   8-bit
Public Const WAVE_FORMAT_1S08 = &H2&                   '/* 11.025 kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_1M16 = &H4&                   '/* 11.025 kHz, Mono,   16-bit
Public Const WAVE_FORMAT_1S16 = &H8&                   '/* 11.025 kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_2M08 = &H10&                  '/* 22.05  kHz, Mono,   8-bit
Public Const WAVE_FORMAT_2S08 = &H20&                  '/* 22.05  kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_2M16 = &H40&                  '/* 22.05  kHz, Mono,   16-bit
Public Const WAVE_FORMAT_2S16 = &H80&                  '/* 22.05  kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_4M08 = &H100&                 '/* 44.1   kHz, Mono,   8-bit
Public Const WAVE_FORMAT_4S08 = &H200&                 '/* 44.1   kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_4M16 = &H400&                 '/* 44.1   kHz, Mono,   16-bit
Public Const WAVE_FORMAT_4S16 = &H800&                 '/* 44.1   kHz, Stereo, 16-bit

Public Const WAVE_FORMAT_PCM = 1

Public Const WHDR_DONE = &H1&              '/* done bit */
Public Const WHDR_PREPARED = &H2&          '/* set if this header has been prepared */
Public Const WHDR_BEGINLOOP = &H4&         '/* loop start block */
Public Const WHDR_ENDLOOP = &H8&           '/* loop end block */
Public Const WHDR_INQUEUE = &H10&          '/* reserved for driver */

Public Const WIM_OPEN = &H3BE
Public Const WIM_CLOSE = &H3BF
Public Const WIM_DATA = &H3C0

Public Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Public Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Public Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long

Public Declare Function waveInGetNumDevs Lib "winmm" () As Long
Public Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long

Public Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal flags As Long) As Long
Public Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Public Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Public Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Public Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long


Sub IniciarSonido(frm As Form)
    Dim Caps As WAVEINCAPS, Which As Long
   'On Error GoTo Err_Proc

    frm.DevicesBox.Clear
    For Which = 0 To waveInGetNumDevs - 1
        Call waveInGetDevCaps(Which, VarPtr(Caps), Len(Caps))
        'If Caps.Formats And WAVE_FORMAT_1M08 Then   WAVE_FORMAT_4S08
        If Caps.Formats And WAVE_FORMAT_4S08 Then 'Now is 1S08 -- Compruebe que los dispositivos que pueden hacer música de 8-bit 11kHz
            Call frm.DevicesBox.AddItem(StrConv(Caps.ProductName, vbUnicode), Which)
        End If
    Next
    If frm.DevicesBox.ListCount = 0 Then
        MsgBox "Usted no tiene dispositivos de entrada de sonido!", vbCritical, "Ack!"
        End
    End If
    frm.DevicesBox.ListIndex = 0


Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "mdlSonido", "IniciarSonido"
   Err.Clear
   Resume Exit_Proc

End Sub
