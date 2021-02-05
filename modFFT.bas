Attribute VB_Name = "modFFT"
'--------------------------------------------------------------------
' VB FFT Release 2-B
' by Murphy McCauley (MurphyMc@Concentric.NET)
' 10/01/99
'--------------------------------------------------------------------
'Acerca de:
'Este c�digo es muy, muy fuertemente basada en la Cruz de Don fourier.pas
'Turbo Pascal Unidad para el c�lculo de la Transformada R�pida de Fourier.
'Yo no he ejecutado la totalidad de sus funciones, aunque tambi�n puede hacerlo
'Por lo que en el futuro.
'Para m�s informaci�n, puede ponerse en contacto conmigo por correo electr�nico, revisar mi sitio web en:
'http://www.fullspectrum.com/deeth/~~V
'O consulte la p�gina de Don Cruz FFT web en:
'Http://www.intersrv.com/ ~ dcross / fft.html
'Tambi�n es posible que le interese en el FFT.DLL que puse la base, junto
'En la FFT Don Cruz de c�digo C. Es exigible con Visual Basic y
'Incluye VB declara. Usted puede obtener desde cualquier sitio web.
'------------------------------------------------- -------------------
'Historia de la versi�n 2-B:
'Se ha corregido un par de errores que resultaron de m� perder el tiempo con
'Los nombres de las variables despu�s de la implementaci�n y no volver a comprobar. MAL ME.
'--------
'Historia de la Versi�n 2:
'Alta FrequencyOfIndex () que es Index_to_frequency Don Cross ().
'FourierTransform () ahora puede hacer las transformaciones inversas.
'Alta CalcFrequency (), que puede hacer una transformaci�n para una sola
'Frecuencia.
'------------------------------------------------- -------------------
'Uso:
'Las funciones �tiles son:
'FourierTransform () realiza una transformada r�pida de Fourier en un par de
'Las matrices de doble - un real, imaginario. No quieren / necesitan
'De los n�meros imaginarios? S�lo tiene que utilizar una serie de 0s. Esta funci�n puede
'Tambi�n hacemos FFT inversa.
'FrequencyOfIndex () puedo decir lo que la frecuencia real de un determinado �ndice
'Corresponde a.
'CalcFrequency () transforma una sola frecuencia.
'------------------------------------------------- -------------------
'Notas:
'Todos los arreglos deben ser 0 base (es decir, Dim laLista (0 a 1023) o
'Dim laLista (1023)).
'El n�mero de muestras debe ser una potencia de dos (es decir 2 ^ x).
'FrequencyOfIndex () y CalcFrequency () no se han probado mucho.
'Utilice esta BAJO SU PROPIO RIESGO.
'--------------------------------------------------------------------

Option Explicit

'Private Const PI        As Double = 3.14159265358979

Private m_lngPowers(16) As Long

Private Function NumberOfBitsNeeded(ByVal PowerOfTwo As Long) As Long
    Dim i               As Long
    
    If m_lngPowers(0) = 0 Then
        For i = 0 To UBound(m_lngPowers)
            m_lngPowers(i) = 2 ^ i
        Next
    End If

    For i = 0 To 16
        If (PowerOfTwo And m_lngPowers(i)) <> 0 Then
            NumberOfBitsNeeded = i
            Exit Function
        End If
    Next
End Function

Private Function IsPowerOfTwo(ByVal X As Long) As Boolean
    If (X < 2) Then Exit Function
    IsPowerOfTwo = Not (X And (X - 1))
End Function

Private Function ReverseBits(ByVal Index As Long, ByVal NumBits As Long) As Long
    Dim i As Long, Rev As Long

    For i = 0 To NumBits - 1
        Rev = (Rev * 2) Or (Index And 1)
        Index = Index \ 2
    Next

    ReverseBits = Rev
End Function

Public Sub FourierTransform( _
    ByVal NumSamples As Long, _
    RealIn() As Double, _
    ImageIn() As Double, _
    RealOut() As Double, _
    ImagOut() As Double, _
    Optional InverseTransform As Boolean = False _
)

    Dim i           As Long, j          As Long
    Dim k           As Long, n          As Long

    Dim BlockSize   As Long, BlockEnd   As Long

    Dim DeltaAngle  As Double, DeltaAr  As Double
    Dim Alpha       As Double, Beta     As Double

    Dim TR          As Double, TI       As Double
    Dim AR          As Double, AI       As Double
    
    Dim AngleNumerator                  As Double

    Dim NumBits                         As Long

    If InverseTransform Then
        AngleNumerator = -2# * PI
    Else
        AngleNumerator = 2# * PI
    End If

    NumBits = NumberOfBitsNeeded(NumSamples)

    For i = 0 To (NumSamples - 1)
        j = ReverseBits(i, NumBits)
        RealOut(j) = RealIn(i)
        ImagOut(j) = ImageIn(i)
    Next

    BlockEnd = 1
    BlockSize = 2

    Do While BlockSize <= NumSamples
        DeltaAngle = AngleNumerator / BlockSize
        Alpha = Sin(0.5 * DeltaAngle)
        Alpha = 2# * Alpha * Alpha
        Beta = Sin(DeltaAngle)

        i = 0
        Do While i < NumSamples
            AR = 1#
            AI = 0#
            
            j = i
            For n = 0 To BlockEnd - 1
                k = j + BlockEnd
                TR = AR * RealOut(k) - AI * ImagOut(k)
                TI = AI * RealOut(k) + AR * ImagOut(k)
                RealOut(k) = RealOut(j) - TR
                ImagOut(k) = ImagOut(j) - TI
                RealOut(j) = RealOut(j) + TR
                ImagOut(j) = ImagOut(j) + TI
                DeltaAr = Alpha * AR + Beta * AI
                AI = AI - (Alpha * AI - Beta * AR)
                AR = AR - DeltaAr
                j = j + 1
            Next

            i = i + BlockSize
        Loop

        BlockEnd = BlockSize
        BlockSize = BlockSize * 2
    Loop

    If InverseTransform Then
        For i = 0 To NumSamples - 1
            RealOut(i) = RealOut(i) / NumSamples
            ImagOut(i) = ImagOut(i) / NumSamples
        Next
    End If
End Sub
