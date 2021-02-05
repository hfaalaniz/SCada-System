Attribute VB_Name = "mdlUno"
Option Explicit

Public a As Double
Public b As Double
Public c As Double
Public ana As Double
Public AnB As Double
Public anc As Double

Public Sub rightab(a, b)

    c = Sqr((a * a) + (b * b))
    anc = (Atn((a / b))) * (180 / PI)
    AnB = (Atn((b / a))) * (180 / PI)

End Sub

Public Sub rightac(a, c)

    b = Sqr((c * c) - (a * a))
    AnB = (Atn((b / a))) * (180 / PI)
    anc = (Atn((a / b))) * (180 / PI)

End Sub

Public Sub rightbc(b, c)

    a = Sqr((c * c) - (b * b))
    AnB = (Atn((b / a))) * (180 / PI)
    anc = (Atn((a / b))) * (180 / PI)

End Sub

Public Sub rightangba(a, AnB)

    anc = 90 - AnB
    c = a / Sin((anc * (PI / 180)))
    b = Sqr((c * c) - (a * a))

End Sub

Public Sub rightangbc(c, AnB)

    anc = 90 - AnB
    a = Sin((anc * (PI / 180))) * c
    b = Sqr((c * c) - (a * a))

End Sub

Public Sub rightangca(a, anc)

    AnB = 90 - anc
    c = a / Sin((anc * (PI / 180)))
    b = Sqr((c * c) - (a * a))

End Sub

Public Sub rightangcc(c, anc)

    AnB = 90 - anc
    a = Sin((anc * (PI / 180))) * c
    b = Sqr((c * c) - (a * a))

End Sub

Public Sub rightangcb(b, anc)

    AnB = 90 - anc
    a = b * Tan(anc * (PI / 180))
    c = Sqr((b * b) + (a * a))

End Sub

Public Sub rightangbb(b, AnB)

    anc = 90 - AnB
    a = b * Tan(anc * (PI / 180))
    c = Sqr((b * b) + (a * a))

End Sub


