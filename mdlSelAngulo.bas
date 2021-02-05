Attribute VB_Name = "mdlSelAngulo"
Option Explicit

Public Type Coords 'Handles Coordinate Systems
X As Long
Y As Long
End Type

Public InterpretedMouse As Coords 'Where the program raycasts the mouse position
Public theAngle As Double         'Variable that holds the angle

Private Midlpoint As Coords       'Middlepoint of the circle
Private Mouse As Coords           'Mouse Coordinates
Private Source As Coords          'Top of circle (0 degrees)
Private Source2 As Coords         'Right of circle (270 degrees)


Private LinA As Long              'These hold the various lengths of lines
Private LinB As Long
Private LinC As Long
Private LinD As Long
Private LinF As Long

Private TheSine                   'These hold the all important raycasting numbers
Private TheCoSine

Private IsDoing As Boolean        'Prevent Annoying Problems

Public Sub DibujarAngulo(pic As PictureBox, X As Single, Y As Single)
If IsDoing Then Exit Sub
IsDoing = True

'It helps if the picture box is a square
pic.Cls
pic.Circle (pic.ScaleWidth / 2, pic.ScaleHeight / 2), ((((pic.ScaleWidth / 2) + (pic.ScaleHeight / 2)) / 2) - 50), vbBlack 'Draw Circle

DoEvents

Mouse.X = X       'Set Mouse
Mouse.Y = Y

Midlpoint.X = pic.ScaleWidth / 2 'Set Middlepoint
Midlpoint.Y = pic.ScaleHeight / 2

Source.X = pic.ScaleWidth / 2    'Set 0 Degrees
Source.Y = (pic.ScaleHeight / 2) - ((((pic.ScaleWidth / 2) + (pic.ScaleHeight / 2)) / 2) - 50)

Source2.X = pic.ScaleWidth - 100 'set 270 degrees
Source2.Y = pic.ScaleHeight / 2

If Mouse.X > Source2.X Then Mouse.X = Source2.X
If Mouse.X < 100 Then Mouse.X = 100

LinA = Sqr(((Mouse.X - Midlpoint.X) ^ 2) + ((Mouse.Y - Midlpoint.Y) ^ 2))
LinB = Sqr(((Source.X - Midlpoint.X) ^ 2) + ((Source.Y - Midlpoint.Y) ^ 2))
LinC = Sqr(((Mouse.X - Source.X) ^ 2) + ((Mouse.Y - Source.Y) ^ 2))



'pic.Line (Midlpoint.X, Midlpoint.Y)-(Mouse.X, Mouse.Y), RGB(255, 0, 0)
'pic.Line (Midlpoint.X, Midlpoint.Y)-(Source.X, Source.Y), RGB(0, 255, 0)
'pic.Line (Source.X, Source.Y)-(Mouse.X, Mouse.Y), RGB(0, 0, 255)
'pic.Line (Midlpoint.X, Midlpoint.Y)-((Source.X + Mouse.X) / 2, (Source.Y + Mouse.Y / 2))
'pic.Line (Source2.X, Source2.Y)-(Mouse.X, Mouse.Y), RGB(0, 0, 255)
'pic.Line (Source2.X, Source2.Y)-(Midlpoint.X, Midlpoint.Y), RGB(0, 255, 0)
'pic.Line (Midlpoint.X, Midlpoint.Y)-((Mouse.X + Source2.X) / 2, (Mouse.Y + Source2.Y) / 2)

TheSine = Sin(DegtoRad(270 - theAngle)) * (Source2.X - Midlpoint.X)
TheCoSine = Cos(DegtoRad(270 - theAngle)) * (Source2.X - Midlpoint.X)

pic.Line (Midlpoint.X, Midlpoint.Y)-(Midlpoint.X + TheCoSine, Midlpoint.Y + TheSine)

InterpretedMouse.X = Midlpoint.X + TheCoSine
InterpretedMouse.Y = Midlpoint.Y + TheSine

LinD = Sqr(((Mouse.X - Source2.X) ^ 2) + ((Mouse.Y - Source2.Y) ^ 2))
LinF = Sqr(((Midlpoint.X - Source2.X) ^ 2) + ((Midlpoint.Y - Source2.Y) ^ 2))

If Mouse.Y > pic.ScaleHeight / 2 And Mouse.X > Midlpoint.X / 2 Then
        theAngle = Abs((2 * RadtoDeg(ArcSine((LinD / 2) / LinF))) - 270)
        GoTo OverSelectCase
End If
Select Case Mouse.X
    Case Is < Midlpoint.X
        theAngle = 2 * (RadtoDeg(ArcSine((LinC / 2) / LinB)))
    Case Is > Midlpoint.X
        theAngle = 360 - (2 * (RadtoDeg(ArcSine((LinC / 2) / LinB))))
    Case Else
        theAngle = 180
End Select

OverSelectCase:
'Debug.Print LinA
'Debug.Print LinB
'Debug.Print LinC
'Debug.Print vbCrLf
'Debug.Print LinD
'Debug.Print LinF
'Debug.Print vbCrLf
IsDoing = False
End Sub
Private Function RadtoDeg(TheNumber As Double) As Double
RadtoDeg = (TheNumber * 180) / 3.1415926535
End Function
Private Function DegtoRad(TheNumber As Double) As Double
DegtoRad = ((TheNumber * 3.1415926535) / 180)
End Function
Private Function ArcSine(TheNumber As Double) As Double
On Error Resume Next
ArcSine = Atn(TheNumber / Sqr(-TheNumber * TheNumber + 1))
End Function

