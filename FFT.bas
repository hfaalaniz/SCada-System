Attribute VB_Name = "FFT"
Option Explicit

' Radix-2 FFT by Murphy McCauley

Private Const PI        As Double = 3.14159265358979
Private Const PI2       As Double = PI * 2

Private m_lngP2(16)     As Long
Private m_sngDLA(16)    As Single
Private m_sngDLB(16)    As Single

Public Sub InitFFT()
    Dim i As Long
    
    ' Lookup Tables for Alpha and Beta
   'On Error GoTo Err_Proc

    If m_lngP2(0) = 0 Then
        For i = 0 To 16
            m_lngP2(i) = 2 ^ i
            m_sngDLA(i) = 2 * Sin(0.5 * PI2 / (m_lngP2(i) * 2)) ^ 2
            m_sngDLB(i) = Sin(PI2 / (m_lngP2(i) * 2))
        Next
    End If

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "FFT", "InitFFT"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub RealFFT( _
    ByVal NumMuestras As Long, _
          EntradaReal() As Single, _
          SalidaReal() As Single, _
          SalidaImag() As Single)

    Static Rev         As Long, NumBits    As Long

    Static i           As Long, j          As Long
    Static k           As Long, N          As Long
    Static L           As Long

    Static BlockSize   As Long, BlockEnd   As Long

    Static DeltaAr     As Single
    Static Alpha       As Single, Beta     As Single

    Static TR          As Single, TI       As Single
    Static AR          As Single, AI       As Single

   ''On Error GoTo Err_Proc

    For N = 0& To 16&
        If NumMuestras = m_lngP2(N) Then
            NumBits = N
            Exit For
        End If
    Next
    
    For i = 0& To NumMuestras - 1&
        Rev = 0&
        k = i

        For j = 0& To NumBits - 1&
            Rev = (Rev * 2&) Or (k And 1&)
            k = k \ 2&
        Next

        SalidaReal(Rev) = EntradaReal(i)
    Next

    BlockEnd = 1
    BlockSize = 2
    L = 0

    Do While BlockSize <= NumMuestras
        Alpha = m_sngDLA(L)
        Beta = m_sngDLB(L)
        L = (L + 1) Mod NumBits

        For i = 0& To NumMuestras - 1 Step BlockSize
            AR = 1#
            AI = 0#
            
            j = i
            For N = 0& To BlockEnd - 1&
                k = j + BlockEnd
                TR = AR * SalidaReal(k) - AI * SalidaImag(k)
                TI = AI * SalidaReal(k) + AR * SalidaImag(k)
                SalidaReal(k) = SalidaReal(j) - TR
                SalidaImag(k) = SalidaImag(j) - TI
                SalidaReal(j) = SalidaReal(j) + TR
                SalidaImag(j) = SalidaImag(j) + TI
                DeltaAr = Alpha * AR + Beta * AI
                AI = AI - Alpha * AI + Beta * AR
                AR = AR - DeltaAr
                j = j + 1&
            Next
        Next

        BlockEnd = BlockSize
        BlockSize = BlockSize * 2
    Loop

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "FFT", "RealFFT"
   Err.Clear
   Resume Exit_Proc

End Sub
