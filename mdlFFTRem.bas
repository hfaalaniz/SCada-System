Attribute VB_Name = "mdlFFTRem"
Global REX(2048)  'REX[ ] holds the real part of the frequency domain
Global IMX(2048) 'IMX[ ] holds the imaginary part of the frequency domain
Global Const n = 2048

Public Sub fft()

'PI = 3.14159265 'Set constants

1000 'THE FAST FOURIER TRANSFORM
'copyright © 1997-1999 by California Technical Publishing
'published with  permission from Steven W Smith, www.dspguide.com
'GUI by logix4u , www.logix4u.net
'modified by logix4u, www.logix4.net
1010 'Upon entry, N% contains the number of points in the DFT, REX[ ] and
1020 'IMX[ ] contain the real and imaginary parts of the input. Upon return,
1030 'REX[ ] and IMX[ ] contain the DFT output. All signals run from 0 to N%-1.
1060 NM1% = n% - 1
1070 ND2% = n% / 2
1080 m% = CInt(Log(n%) / Log(2))
1090 j% = ND2%
1100 '
1110 For i% = 1 To n% - 2 'Bit reversal sorting
1120    If i% >= j% Then GoTo 1190
1130    TR = REX(j%)
1140    TI = IMX(j%)
1150    REX(j%) = REX(i%)
1160    IMX(j%) = IMX(i%)
1170    REX(i%) = TR
1180    IMX(i%) = TI
1190    k% = ND2%
1200    If k% > j% Then GoTo 1240
1210    j% = j% - k%
1220    k% = k% / 2
1230    GoTo 1200
1240    j% = j% + k%
1250 Next i%
1260 '
1270 For L% = 1 To m% 'Loop for each stage
1280    LE% = CInt(2 ^ L%)
1290    LE2% = LE% / 2
1300    UR = 1
1310    UI = 0
1320    SR = Cos(PI / LE2%) 'Calculate sine & cosine values
1330    SI = -Sin(PI / LE2%)
1340    For j% = 1 To LE2% 'Loop for each sub DFT
1350        JM1% = j% - 1
1360        For i% = JM1% To NM1% Step LE% 'Loop for each butterfly
1370            IP% = i% + LE2%
1380            TR = REX(IP%) * UR - IMX(IP%) * UI 'Butterfly calculation
1390            TI = REX(IP%) * UI + IMX(IP%) * UR
1400            REX(IP%) = REX(i%) - TR
1410            IMX(IP%) = IMX(i%) - TI
1420            REX(i%) = REX(i%) + TR
1430            IMX(i%) = IMX(i%) + TI
1440        Next i%
1450        TR = UR
1460        UR = TR * SR - UI * SI
1470        UI = TR * SI + UI * SR
1480    Next j%
1490 Next L%
1500 '
End Sub


