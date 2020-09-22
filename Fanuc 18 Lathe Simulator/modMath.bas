Attribute VB_Name = "modMath"
Option Explicit

Private Type CIRCLEMIDDLEPOINT
    x As Double
    y As Double
End Type

Private Type RADIUSDEGREES
    Start As String
    End As String
End Type

Private Const PI As Double = 3.14159265358979  ' 4 * Atn(1)

Public CircleCenter As CIRCLEMIDDLEPOINT
Public CircleDegrees As RADIUSDEGREES

Private Function ArcCos(x As Double) As Double
    If x = 1 Then ArcCos = 0: Exit Function
    ArcCos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function

Private Function LineLen(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
    Dim A As Double, B As Double
    A = Abs(X2 - X1)
    B = Abs(Y2 - Y1)
    LineLen = Sqr(A ^ 2 + B ^ 2)
End Function

Private Sub CalcAnAngle(CenterX As Double, CenterY As Double, r As Double, xNull As Double, yNull As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, ClockWise As Boolean)
    On Error Resume Next
    Dim dDegrees As Double
    Dim SideA As Double, SideB As Double, SideC As Double
    Dim A As Double
    SideC = LineLen(CenterX, CenterY, xNull, yNull)
    SideB = LineLen(CenterX, CenterY, X3, Y3)
    SideA = LineLen(X3, Y3, xNull, yNull)
    A = ArcCos((SideA ^ 2 - SideB ^ 2 - SideC ^ 2) / (SideB * SideC * -2))
    A = A * (180 / 3.141)
    If Round(X3, 3) < Round(xNull, 3) Then
        dDegrees = A + 180
    ElseIf Round(X3, 3) = Round(xNull, 3) Then
        If Round(Y3, 3) < Round(CenterY, 3) Then
            dDegrees = 180
        Else
            dDegrees = 0
        End If
    ElseIf Round(X3, 3) > Round(xNull, 3) Then
        dDegrees = 180 - A
    Else
        dDegrees = A
    End If
    CircleDegrees.Start = dDegrees
    SideC = LineLen(CenterX, CenterY, xNull, yNull)
    SideB = LineLen(CenterX, CenterY, X4, Y4)
    SideA = LineLen(X4, Y4, xNull, yNull)
    A = ArcCos((SideA ^ 2 - SideB ^ 2 - SideC ^ 2) / (SideB * SideC * -2))
    A = A * (180 / 3.141)
    If Round(X4, 3) < Round(xNull, 3) Then
        dDegrees = A + 180
    ElseIf Round(X4, 3) = Round(xNull, 3) Then
        If Round(Y4, 3) < Round(CenterY, 3) Then
            dDegrees = 180
        Else
            dDegrees = 0 'A
        End If
    ElseIf Round(X4, 3) > Round(xNull, 3) Then
        dDegrees = 180 - A
    Else
        dDegrees = A
    End If
    CircleDegrees.End = dDegrees
End Sub

Public Sub GetCircleCenter(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Radius As Double, ClockWise As Boolean)
    Dim Dist As Double
    Dim A As Double
    Dim B As Double
    Dim C As Double
    On Error Resume Next
    'ClockWise = True
    Dist = Sqr((Radius ^ 2) - (Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2)) / 2) ^ 2)
    A = 1 / 2 * (X1 + X2)
    B = (Y2 - Y1) * Dist
    C = Sqr(((Y2 - Y1) ^ 2) + ((X2 - X1) ^ 2))
    If ClockWise = True Then
        CircleCenter.x = Round(A + (B / C), 3)
    Else
        CircleCenter.x = Round(A - (B / C), 3)
    End If
    CircleCenter.x = Int(CircleCenter.x + 0.5)
    A = 1 / 2 * (Y1 + Y2)
    B = (X2 - X1) * Dist
    If ClockWise = True Then
        CircleCenter.y = Round(A - (B / C), 3)
    Else
        CircleCenter.y = Round(A + (B / C), 3)
    End If
    CircleCenter.y = Int(CircleCenter.y + 0.5)
End Sub

Public Sub GetCircleDegrees(xCenter As Double, yCenter As Double, r As Double, xSPoint As Double, ySPoint As Double, xEPoint As Double, yEpoint As Double, ClockWise As Boolean)
    Dim xNull As Double
    Dim yNull As Double
    xNull = xCenter
    yNull = yCenter - r
    CalcAnAngle xCenter, yCenter, r, xNull, yNull, xSPoint, ySPoint, xEPoint, yEpoint, ClockWise
End Sub

