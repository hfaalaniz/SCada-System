Attribute VB_Name = "mdlFFTCompleja"
Option Explicit
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Type Complex
  Re As Double
  Im As Double
End Type
Public Const PI = 3.1415926535
Public TD(1023) As Complex, FD(1023) As Complex
Public y(1023) As Double

Public Function Complex_Compue(a As Complex, b As Complex, Mark As Byte) As Complex
  '????
  Select Case Mark
    Case 1 '???
      Complex_Compue.Re = a.Re + b.Re
      Complex_Compue.Im = a.Im + b.Im
    Case 2 '???
      Complex_Compue.Re = a.Re - b.Re
      Complex_Compue.Im = a.Im - b.Im
    Case 3 '???
      Complex_Compue.Re = a.Re * b.Re - a.Im * b.Im
      Complex_Compue.Im = a.Re * b.Im + a.Im * b.Re
  End Select
End Function

Public Sub FFT(TD() As Complex, FD() As Complex, m As Long)
  Dim i As Integer, j As Integer, k As Integer, bfSize As Integer, p As Integer
  Dim Angle As Double
  Dim N As Long
  N = 2 ^ m
  ReDim x1(N - 1) As Complex, x2(N - 1) As Complex, X(N - 1) As Complex, W(N \ 2 - 1) As Complex ', TD(n - 1) As Complex, FD(n - 1) As Complex
  For i = 0 To N \ 2 - 1
    Angle = -i * PI * 2 / N
    With W(i)
      .Re = Cos(Angle)
      .Im = Sin(Angle)
    End With
  Next i
  CopyMemory x1(0), TD(0), 16 * N
  For k = 0 To m - 1 '???????
    For j = 0 To 2 ^ k - 1
      bfSize = 2 ^ (m - k)
      For i = 0 To bfSize \ 2 - 1
        p = j * bfSize
        x2(i + p) = Complex_Compue(x1(i + p), x1(i + p + bfSize \ 2), 1)
        x2(i + p + bfSize \ 2) = Complex_Compue(Complex_Compue(x1(i + p), x1(i + p + bfSize \ 2), 2), W(i * (2 ^ k)), 3)
      Next i
    Next j
    '??X1?X2???
    CopyMemory X(0), x1(0), 16 * N
    CopyMemory x1(0), x2(0), 16 * N
    CopyMemory x2(0), X(0), 16 * N
  Next k
  For j = 0 To N - 1 '??
    p = 0
    For i = 0 To m - 1
      If (j And (2 ^ i)) = (2 ^ i) Then p = p + 2 ^ (m - i - 1)
    Next i
    FD(j) = x1(p)
  Next j
End Sub

