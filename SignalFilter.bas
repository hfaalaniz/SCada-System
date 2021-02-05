Attribute VB_Name = "SignalFilter"
Option Explicit


' Windowed Sinc Filtro and Convolution
'Factor: Valor de 0,0 a 0,5. Este es el punto de corte.
'BSP:    Señal muestreda con con 22.050 Hz, la frecuencia de corte  debe ser 220,5 Hz.
'Factor: = 220.5/22050 = 0,01
'Taps: Calidad del filtro. Los grifos más, cuanto más pronunciada
'La banda de atenuación es sin embargo, los grifos más, cuanto más tiempo mas retraso en los calculos
'El retraso, y los cálculos más.
'#################################################################
' compiled with MS VC++ 2005
' int __stdcall Convolve(float * Muestras, int nSamples,
'                        float * kernel,  int nKernel,
'                        float * output,  float * overlap)
' {
'     int i, j, k;
'     for (i=0; i < nSamples; i++)
'         for (j=0; j < nKernel; j++)
'             output[i+j] += Muestras[i] * kernel[j];
'     k = nSamples >= nKernel ? nKernel : nSamples;
'
'     for (i=0; i < k; i++)
'     {
'         Muestras[i] += overlap[i];
'         overlap[i]  = output[nSamples+i];
'     }
'     return nSamples;
' }
'#################################################################
Private Const ASM_CONVOLUTION As String = _
    "8B44240853558B6C241C565785C07E3C8B5C24148BF52BDD894424" & _
    "188B7C242085FF7E1A8B54241C8BCED90433D80A83C20483C1044F" & _
    "D841FCD959FC75EC8B4C241883C60449894C241875D085C07E158B" & _
    "4C24148BD52BD18BF08B3C0A893983C1044E75F58B4C24288B5424" & _
    "1433FF8D7485002BD18B5C24203BC37D028BD83BFB7D18D9040AD8" & _
    "014783C60483C104D95C0AFC8B5EFC8959FCEBDA5F5E5D5BC21800"

Private m_blnFastConv   As Boolean
Private m_udtHook       As HookData
Private m_udtASM        As MachineCode


'#################################################################
'#################################################################


Public Type FilterKernel
    kernel()            As Single
    olap()              As Single
    taps                As Long
End Type

Public Enum FilterType
    FiltroPasaAltos = 0
    FiltroPasaBajos
End Enum

Private Const PI        As Single = 3.14159265358979
Private Const PI2       As Single = PI * 2


' Calcular en la ventana sinc el Filtro FIR del núcleo
' from http://www.dspguide.com/
Public Function CrearFiltro(ByVal ftp As FilterType, ByVal taps As Long, ByVal factor As Single) As FilterKernel
    Dim omega       As Single
    Dim N           As Single
    Dim sum         As Single
    Dim m           As Long
    Dim i           As Long

   'On Error GoTo Err_Proc

    omega = PI2 * factor
    m = taps / 2

    With CrearFiltro
        ReDim .kernel(taps - 1) As Single
        ReDim .olap(taps - 1) As Single
        .taps = taps
    
        ' Sinc function
        For i = 0 To taps - 1
            If i - m = 0 Then
                .kernel(i) = omega
            Else
                N = i - m
                .kernel(i) = Sin(omega * N) / N
            End If
            'minimizar en forma de campana las ondas en la ventana
            .kernel(i) = .kernel(i) * VentanaHamming(i, taps)
            sum = sum + .kernel(i)
        Next
        
        If sum = 0 Then sum = 1
        
        For i = 0 To taps - 1
            ' Normalizar kernel
            .kernel(i) = .kernel(i) / sum
        Next
        
        If ftp = FiltroPasaAltos Then
            ' invertir los dos espectrales inversión al convertir de paso bajo a un paso alto
            For i = 0 To taps - 1
                .kernel(i) = -.kernel(i)
            Next
            .kernel(m) = .kernel(m) + 1
        End If
    End With

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "SignalFilter", "CrearFiltro"
   Err.Clear
   Resume Exit_Proc

End Function


' clears overlap of the Filtro
Public Sub ResetFilter(kernel As FilterKernel)
    Dim i           As Long
    
    For i = 0 To kernel.taps - 1
        kernel.olap(i) = 0
    Next
End Sub


' Filtro a signal with Overlap-Add method
Public Sub FiltrarProceso(sngValues() As Single, kernel As FilterKernel)
    Dim sngOut()    As Single
    Dim i           As Long
    Dim N           As Long
    
   'On Error GoTo Err_Proc

    N = UBound(sngValues) + 1
    
    If m_blnFastConv Then
        ReDim sngOut(N + kernel.taps - 1) As Single
        FastConvolve sngValues(0), N, _
                     kernel.kernel(0), kernel.taps, _
                     sngOut(0), kernel.olap(0)
    Else
        If N >= kernel.taps - 1 Then
            sngOut = Convolve(sngValues, kernel.kernel)
    
            For i = 0 To kernel.taps - 1
                sngOut(i) = sngOut(i) + kernel.olap(i)
                kernel.olap(i) = sngOut(N + i)
            Next
    
            For i = 0 To N - 1
                sngValues(i) = sngOut(i)
            Next
        End If
    End If

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "SignalFilter", "FiltrarProceso"
   Err.Clear
   Resume Exit_Proc

End Sub


' convolution (Input Side method)
Private Function Convolve(a() As Single, B() As Single) As Single()
    Dim c() As Single
    Dim i   As Long
    Dim j   As Long
    
   'On Error GoTo Err_Proc

    ReDim c(UBound(a) + UBound(B) + 1) As Single
    
    For i = 0 To UBound(a)
        For j = 0 To UBound(B)
            c(i + j) = c(i + j) + a(i) * B(j)
        Next
    Next
    
    Convolve = c

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "SignalFilter", "Convolve"
   Err.Clear
   Resume Exit_Proc

End Function


Private Function FastConvolve( _
    Muestras As Single, ByVal nSamples As Long, _
    kernel As Single, ByVal nKernel As Long, _
    output As Single, overlap As Single _
) As Long

    FastConvolve = -1
End Function


Public Sub InitFastConvolution()
   'On Error GoTo Err_Proc

    If Not m_blnFastConv Then
        m_udtASM = ASMStringToMemory(ASM_CONVOLUTION)
        m_udtHook = RedirectFunction(AddressOf FastConvolve, True, m_udtASM.pAsm)
        m_blnFastConv = True
    End If

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "SignalFilter", "InitFastConvolution"
   Err.Clear
   Resume Exit_Proc

End Sub


Public Sub TerminateFastConvolution()
   'On Error GoTo Err_Proc

    If m_blnFastConv Then
        FreeASMMemory m_udtASM
        RestoreFunction m_udtHook
        m_blnFastConv = False
    End If

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "SignalFilter", "TerminateFastConvolution"
   Err.Clear
   Resume Exit_Proc

End Sub


Public Function VentanaHamming(ByVal i As Single, ByVal N As Single) As Single
    VentanaHamming = 0.54 - 0.46 * Cos(PI2 * i / N)
End Function

