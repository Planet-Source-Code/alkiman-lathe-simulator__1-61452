Attribute VB_Name = "modPaint"
Option Explicit

Private Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Const PI As Double = 3.14159265358979  ' 4 * Atn(1)

Private Type Rect
    Left As Double
    Top As Double
    Right As Double
    Bottom As Double
End Type

Private Function CheckDblSpace(sData As String) As String
    Dim sData1 As String
    Dim sChar As String
    Dim iCount As Integer
    sData = UCase(sData)
    sData1 = ""
    For iCount = 1 To Len(sData)
        sChar = Mid(sData, iCount, 1)
        If sChar = "," Then
            sData1 = sData1 & "."
        ElseIf sChar <> " " Then
            sData1 = sData1 & sChar
        End If
    Next
    CheckDblSpace = CheckCommand(sData1)
End Function

Public Sub DrawContour(sData As String, blnRadius As Boolean, DrawPicture As PictureBox, Optional DrawFast As Boolean = False)
    Dim sData1 As String
    Static blnLastCommandG00 As Boolean
    If InStr(1, sData, "G0 ") <> 0 Or InStr(1, sData, "G00") <> 0 Then
        blnLastCommandG00 = True
    ElseIf InStr(1, sData, "G") <> 0 Then
        blnLastCommandG00 = False
    End If
    If blnRadius = False Then
        'Line
        If InStr(1, sData, "X") <> 0 Or InStr(1, sData, "Z") <> 0 Then
            If InStr(1, sData, "X") <> 0 And InStr(1, sData, "Z") <> 0 Then
                xEndPosition = xWidth + GetCoordinates(sData, InStr(1, sData, "Z"))
                yEndPosition = yHeight - (GetCoordinates(sData, InStr(1, sData, "X")) / 2)
                If blnLastCommandG00 = True Then
                    Dim xMoveDiff As Double
                    Dim zMoveDiff As Double
                    If Int(xStartPosition) > Int(xEndPosition) Then
                        zMoveDiff = xStartPosition - xEndPosition
                    ElseIf Int(xEndPosition) > Int(xStartPosition) Then
                        zMoveDiff = xEndPosition - xStartPosition
                    Else
                        zMoveDiff = 0
                    End If
                    If Int(yStartPosition) > Int(yEndPosition) Then
                        xMoveDiff = yStartPosition - yEndPosition
                    ElseIf Int(yEndPosition) > Int(yStartPosition) Then
                        xMoveDiff = yEndPosition - yStartPosition
                    Else
                        xMoveDiff = 0
                    End If
                    Dim MoveDiff As Double
                    If xMoveDiff >= zMoveDiff Then
                        MoveDiff = Int(zMoveDiff)
                    Else
                        MoveDiff = Int(xMoveDiff)
                    End If
                    If xStartPosition > xEndPosition Then
                        zMoveDiff = xStartPosition - MoveDiff
                    Else
                        zMoveDiff = xStartPosition + MoveDiff
                    End If
                    If yStartPosition > yEndPosition Then
                        xMoveDiff = yStartPosition - MoveDiff
                    Else
                        xMoveDiff = yStartPosition + MoveDiff
                    End If
                    DrawPicture.Line (Int(xStartPosition), Int(yStartPosition))-(Int(zMoveDiff), Int(xMoveDiff))
                    xStartPosition = zMoveDiff
                    yStartPosition = xMoveDiff
                    'DrawPicture.Line (xStartPosition, yStartPosition)-(xEndPosition, yEndPosition)
                End If
            ElseIf InStr(1, sData, "Z") <> 0 Then
                xEndPosition = xWidth + GetCoordinates(sData, InStr(1, sData, "Z"))
            Else
                yEndPosition = yHeight - (GetCoordinates(sData, InStr(1, sData, "X")) / 2)
            End If
            If DrawPicture.DrawStyle = 2 Or DrawFast = True Then
                DrawPicture.Line (xStartPosition, yStartPosition)-(xEndPosition, yEndPosition)
            Else
                LineByPixel DrawPicture, xStartPosition, yStartPosition, xEndPosition, yEndPosition, lLineColor
            End If
            xStartPosition = xEndPosition
            yStartPosition = yEndPosition
            Wait lTimeMilliSeconds
        ElseIf InStr(1, sData, "U") <> 0 Or InStr(1, sData, "W") <> 0 Then
            If InStr(1, sData, "U") <> 0 And InStr(1, sData, "W") <> 0 Then
                xEndPosition = xStartPosition + GetCoordinates(sData, InStr(1, sData, "W"))
                yEndPosition = yStartPosition - ((GetCoordinates(sData, InStr(1, sData, "U")) / 2))
            ElseIf InStr(1, sData, "W") <> 0 Then
                xEndPosition = xStartPosition + GetCoordinates(sData, InStr(1, sData, "W"))
            Else
                yEndPosition = yStartPosition - ((GetCoordinates(sData, InStr(1, sData, "U")) / 2))
            End If
            If DrawPicture.DrawStyle = 2 Or DrawFast = True Then
                DrawPicture.Line (xStartPosition, yStartPosition)-(xEndPosition, yEndPosition)
            Else
                LineByPixel DrawPicture, xStartPosition, yStartPosition, xEndPosition, yEndPosition, lLineColor
            End If
            xStartPosition = xEndPosition
            yStartPosition = yEndPosition
            Wait lTimeMilliSeconds
        End If
        'End Line
    Else
        'Arc
        If InStr(1, sData, "U") <> 0 Or InStr(1, sData, "W") <> 0 Then
            Dim dWX As Double
            Dim dUZ As Double
            dWX = ((-xWidth + xStartPosition) / dScaleFactor) + GetCoordinates(sData, InStr(1, sData, "W"), False)
            dUZ = ((yHeight - yStartPosition) / dScaleFactor * 2) + GetCoordinates(sData, InStr(1, sData, "U"), False)
            sData = Left(sData, InStr(1, sData, "U") - 1) & "X" & dUZ & " " & Right(sData, Len(sData) - InStr(1, sData, "U") - LenCoordinates(sData, InStr(1, sData, "U"), False))
            sData = Left(sData, InStr(1, sData, "W") - 1) & "Z" & dWX & " " & Right(sData, Len(sData) - InStr(1, sData, "W") - LenCoordinates(sData, InStr(1, sData, "W"), False))
            sData = CheckDblSpace(sData)
        End If
        If InStr(1, sData, "X") <> 0 Or InStr(1, sData, "Z") <> 0 Then
            Dim dRadius As Double
            Dim dCenterX As Double
            Dim dCenterY As Double
            Dim ArcRectangle As Rect
            If InStr(1, sData, "R") = 0 Then
                'Radius excl
                xEndPosition = xWidth + GetCoordinates(sData, InStr(1, sData, "Z"))
                yEndPosition = yHeight - (GetCoordinates(sData, InStr(1, sData, "X")) / 2)
                If GetCoordinates(sData, InStr(1, sData, "K")) = 0 Then
                    dRadius = Abs(GetCoordinates(sData, InStr(1, sData, "I")))
                ElseIf GetCoordinates(sData, InStr(1, sData, "I")) = 0 Then
                    dRadius = Abs(GetCoordinates(sData, InStr(1, sData, "K")))
                Else
                    dRadius = Sqr((Abs(GetCoordinates(sData, InStr(1, sData, "I"))) ^ 2) + (Abs(GetCoordinates(sData, InStr(1, sData, "K"))) ^ 2))
                End If
                dRadius = Round(dRadius, 0)
                If InStr(1, sData, "G02") <> 0 Then
                    'ClockWise Radius excl
                    dCenterX = xStartPosition + GetCoordinates(sData, InStr(1, sData, "K"))
                    dCenterY = yStartPosition - GetCoordinates(sData, InStr(1, sData, "I"))
                    ArcRectangle.Left = (dCenterX - dRadius)
                    ArcRectangle.Top = (dCenterY - dRadius)
                    ArcRectangle.Right = (dCenterX + dRadius)
                    ArcRectangle.Bottom = (dCenterY + dRadius)
                    'Rectangle DrawPicture.hdc, ArcRectangle.Left, ArcRectangle.Top, ArcRectangle.Right, ArcRectangle.Bottom
                    If DrawFast = True Then
                        Arc DrawPicture.hdc, ArcRectangle.Left, ArcRectangle.Top, ArcRectangle.Right, ArcRectangle.Bottom, xEndPosition, yEndPosition, xStartPosition, yStartPosition
                    Else
                        GetCircleDegrees dCenterX, dCenterY, dRadius, xStartPosition, yStartPosition, xEndPosition, yEndPosition, True
                        DrawRadius DrawPicture, dCenterX, dCenterY, dRadius, True, lLineColor, CircleDegrees.Start, CircleDegrees.End
                    End If
                    xStartPosition = xEndPosition
                    yStartPosition = yEndPosition
                Else
                    'Contra ClockWise Radius excl
                    dCenterX = xStartPosition + GetCoordinates(sData, InStr(1, sData, "K"))
                    dCenterY = yStartPosition - GetCoordinates(sData, InStr(1, sData, "I"))
                    ArcRectangle.Left = (dCenterX - dRadius)
                    ArcRectangle.Top = (dCenterY - dRadius)
                    ArcRectangle.Right = (dCenterX + dRadius)
                    ArcRectangle.Bottom = (dCenterY + dRadius)
                    'Rectangle DrawPicture.hdc, ArcRectangle.Left, ArcRectangle.Top, ArcRectangle.Right, ArcRectangle.Bottom
                    If DrawFast = True Then
                        Arc DrawPicture.hdc, ArcRectangle.Left, ArcRectangle.Top, ArcRectangle.Right, ArcRectangle.Bottom, xStartPosition, yStartPosition, xEndPosition, yEndPosition
                    Else
                        GetCircleDegrees dCenterX, dCenterY, dRadius, xStartPosition, yStartPosition, xEndPosition, yEndPosition, False
                        DrawRadius DrawPicture, dCenterX, dCenterY, dRadius, False, lLineColor, CircleDegrees.Start, CircleDegrees.End
                    End If
                    xStartPosition = xEndPosition
                    yStartPosition = yEndPosition
                End If
            Else
                'Radius incl
                dRadius = GetCoordinates(sData, InStr(1, sData, "R"))
                If InStr(1, sData, "G02") <> 0 Then
                    'Clockwise Radius incl
                    xEndPosition = xWidth + GetCoordinates(sData, InStr(1, sData, "Z"))
                    yEndPosition = yHeight - (GetCoordinates(sData, InStr(1, sData, "X")) / 2)
                    If xStartPosition = xEndPosition Then
                        GetCircleCenter xStartPosition, yStartPosition, xEndPosition, yEndPosition, dRadius, False 'True or False make no difference
                        ArcRectangle.Left = (CircleCenter.x + dRadius)
                        ArcRectangle.Top = (CircleCenter.y - dRadius)
                        ArcRectangle.Right = (CircleCenter.x - dRadius)
                        ArcRectangle.Bottom = (CircleCenter.y + dRadius)
                        'Rectangle DrawPicture.hdc, ArcRectangle.Left, ArcRectangle.Top, ArcRectangle.Right, ArcRectangle.Bottom
                        If DrawFast = True Then
                            Arc DrawPicture.hdc, ArcRectangle.Left, ArcRectangle.Top, ArcRectangle.Right, ArcRectangle.Bottom, xEndPosition, yEndPosition, xStartPosition, yStartPosition
                        Else
                            GetCircleDegrees CircleCenter.x, CircleCenter.y, dRadius, xStartPosition, yStartPosition, xEndPosition, yEndPosition, True
                            DrawRadius DrawPicture, CircleCenter.x, CircleCenter.y, dRadius, True, lLineColor, CircleDegrees.Start, CircleDegrees.End
                        End If
                    ElseIf yStartPosition = yEndPosition Then
                        GetCircleCenter xStartPosition, yStartPosition, xEndPosition, yEndPosition, dRadius, False 'True or False make no difference
                        ArcRectangle.Left = (CircleCenter.x + dRadius)
                        ArcRectangle.Top = (CircleCenter.y - dRadius)
                        ArcRectangle.Right = (CircleCenter.x - dRadius)
                        ArcRectangle.Bottom = (CircleCenter.y + dRadius)
                        'Rectangle DrawPicture.hdc, ArcRectangle.Left, ArcRectangle.Top, ArcRectangle.Right, ArcRectangle.Bottom
                        If DrawFast = True Then
                            Arc DrawPicture.hdc, ArcRectangle.Left, ArcRectangle.Top, ArcRectangle.Right, ArcRectangle.Bottom, xEndPosition, yEndPosition, xStartPosition, yStartPosition
                        Else
                            GetCircleDegrees CircleCenter.x, CircleCenter.y, dRadius, xStartPosition, yStartPosition, xEndPosition, yEndPosition, True
                            DrawRadius DrawPicture, CircleCenter.x, CircleCenter.y, dRadius, True, lLineColor, CircleDegrees.Start, CircleDegrees.End
                        End If
                    Else
                        GetCircleCenter xStartPosition, yStartPosition, xEndPosition, yEndPosition, dRadius, False
                        ArcRectangle.Left = (CircleCenter.x + dRadius)
                        ArcRectangle.Top = (CircleCenter.y - dRadius)
                        ArcRectangle.Right = (CircleCenter.x - dRadius)
                        ArcRectangle.Bottom = (CircleCenter.y + dRadius)
                        'Rectangle DrawPicture.hdc, ArcRectangle.Left, ArcRectangle.Top, ArcRectangle.Right, ArcRectangle.Bottom
                        If DrawFast = True Then
                            Arc DrawPicture.hdc, ArcRectangle.Left, ArcRectangle.Top, ArcRectangle.Right, ArcRectangle.Bottom, xEndPosition, yEndPosition, xStartPosition, yStartPosition
                        Else
                            GetCircleDegrees CircleCenter.x, CircleCenter.y, dRadius, xStartPosition, yStartPosition, xEndPosition, yEndPosition, True
                            DrawRadius DrawPicture, CircleCenter.x, CircleCenter.y, dRadius, True, lLineColor, CircleDegrees.Start, CircleDegrees.End
                        End If
                    End If
                    xStartPosition = xEndPosition
                    yStartPosition = yEndPosition
                    Wait lTimeMilliSeconds
                Else
                    'Contra Clockwise Radius incl
                    xEndPosition = xWidth + GetCoordinates(sData, InStr(1, sData, "Z"))
                    yEndPosition = yHeight - (GetCoordinates(sData, InStr(1, sData, "X")) / 2)
                    If xStartPosition = xEndPosition Then
                        GetCircleCenter xStartPosition, yStartPosition, xEndPosition, yEndPosition, dRadius, False 'True or False make no difference
                        If CircleCenter.x > xStartPosition Then
                            GetCircleCenter xStartPosition, yStartPosition, xEndPosition, yEndPosition, dRadius, True
                        End If
                        ArcRectangle.Left = (CircleCenter.x + dRadius)
                        ArcRectangle.Top = (CircleCenter.y - dRadius)
                        ArcRectangle.Right = (CircleCenter.x - dRadius)
                        ArcRectangle.Bottom = (CircleCenter.y + dRadius)
                    ElseIf yStartPosition = yEndPosition Then
                        GetCircleCenter xStartPosition, yStartPosition, xEndPosition, yEndPosition, dRadius, False 'True or False make no difference
                        If CircleCenter.y < yStartPosition Then
                            GetCircleCenter xStartPosition, yStartPosition, xEndPosition, yEndPosition, dRadius, True
                        End If
                        ArcRectangle.Left = (CircleCenter.x + dRadius)
                        ArcRectangle.Top = (CircleCenter.y - dRadius)
                        ArcRectangle.Right = (CircleCenter.x - dRadius)
                        ArcRectangle.Bottom = (CircleCenter.y + dRadius)
                    Else
                        GetCircleCenter xStartPosition, yStartPosition, xEndPosition, yEndPosition, dRadius, True
                        ArcRectangle.Left = (CircleCenter.x + dRadius)
                        ArcRectangle.Top = (CircleCenter.y - dRadius)
                        ArcRectangle.Right = (CircleCenter.x - dRadius)
                        ArcRectangle.Bottom = (CircleCenter.y + dRadius)
                    End If
                    'Rectangle DrawPicture.hdc, ArcRectangle.Left, ArcRectangle.Top, ArcRectangle.Right, ArcRectangle.Bottom
                    If DrawFast = True Then
                        Arc DrawPicture.hdc, ArcRectangle.Left, ArcRectangle.Top, ArcRectangle.Right, ArcRectangle.Bottom, xStartPosition, yStartPosition, xEndPosition, yEndPosition
                    Else
                        GetCircleDegrees CircleCenter.x, CircleCenter.y, dRadius, xStartPosition, yStartPosition, xEndPosition, yEndPosition, False
                        DrawRadius DrawPicture, CircleCenter.x, CircleCenter.y, dRadius, False, lLineColor, CircleDegrees.Start, CircleDegrees.End
                    End If
                    xStartPosition = xEndPosition
                    yStartPosition = yEndPosition
                    Wait lTimeMilliSeconds
                End If
            End If
        End If
    End If
End Sub

Private Sub DrawRadius(DrawPicture As PictureBox, xCenter As Double, yCenter As Double, r As Double, ClockWise As Boolean, Color As Long, Optional StartDegrees As String, Optional EndDegrees As String, Optional PassCross As Boolean)
    Dim A As Integer
    Dim n As Double
    Dim xx As Double
    Dim yy As Double
    Dim oldxx As Double
    Dim oldyy As Double
    Dim blnLine As Boolean
    Dim blnNotFirstTime As Boolean
    'PassCross = True
    If StartDegrees = "" Then
        StartDegrees = 0
    End If
    If EndDegrees = "" Then
        EndDegrees = 360
    End If
    If ClockWise = True Then
        If Val(StartDegrees) < Val(EndDegrees) Then
            StartDegrees = Val(StartDegrees) + 360
        End If
    Else
        If Val(EndDegrees) < Val(StartDegrees) Then
            EndDegrees = Val(EndDegrees) + 360
        End If
    End If
    If PassCross = True Then
        DrawPicture.Line (Val(xCenter), Val(yCenter) - 5)-(Val(xCenter), Val(yCenter) + 5), Color
        DrawPicture.Line (Val(xCenter) - 5, Val(yCenter))-(Val(xCenter) + 5, Val(yCenter)), Color
        DrawPicture.Refresh
    End If
    If ClockWise = True Then
        For A = IIf(ClockWise = True, Val(StartDegrees), Val(EndDegrees)) To IIf(ClockWise = True, Val(EndDegrees), Val(StartDegrees)) Step IIf(ClockWise = True, -1, 1)
            n = A * (PI / 180)
            xx = CSng(Sin(n) * Val(r))
            yy = CSng(Cos(n) * Val(r))
            SetPixel DrawPicture.hdc, xx + Val(xCenter), yy + Val(yCenter), Color
            xx = Round(xx, 0)
            yy = Round(yy, 0)
            blnLine = False
            If Abs(oldxx - xx) > 1 Then
                blnLine = True
            End If
            If Abs(oldyy - yy) > 1 Then
                blnLine = True
            End If
            If blnNotFirstTime = False Then
                blnNotFirstTime = True
                blnLine = False
            End If
            If blnLine = True Then
                DrawPicture.Line (oldxx + Val(xCenter), oldyy + Val(yCenter))-(xx + Val(xCenter), yy + Val(yCenter)), Color
            End If
            oldxx = Round(xx, 0)
            oldyy = Round(yy, 0)
            DrawPicture.Refresh
            Wait lTimeMilliSeconds
        Next
    Else
        For A = IIf(ClockWise = True, Val(EndDegrees), Val(StartDegrees)) To IIf(ClockWise = True, Val(StartDegrees), Val(EndDegrees)) Step IIf(ClockWise = True, -1, 1)
            n = A * (PI / 180)
            xx = CSng(Sin(n) * Val(r))
            yy = CSng(Cos(n) * Val(r))
            SetPixel DrawPicture.hdc, xx + Val(xCenter), yy + Val(yCenter), Color
            xx = Round(xx, 0)
            yy = Round(yy, 0)
            blnLine = False
            If Abs(oldxx - xx) > 1 Then
                blnLine = True
            End If
            If Abs(oldyy - yy) > 1 Then
                blnLine = True
            End If
            If blnNotFirstTime = False Then
                blnNotFirstTime = True
                blnLine = False
            End If
            If blnLine = True Then
                DrawPicture.Line (oldxx + Val(xCenter), oldyy + Val(yCenter))-(xx + Val(xCenter), yy + Val(yCenter)), Color
            End If
            oldxx = Round(xx, 0)
            oldyy = Round(yy, 0)
            DrawPicture.Refresh
            Wait lTimeMilliSeconds
        Next
    End If
End Sub

Public Sub DrawXZ_Axis(DrawPicture As PictureBox)
    DrawPicture.Cls
    DrawPicture.ScaleMode = 3
    DrawPicture.ForeColor = vbBlue
    DrawPicture.DrawStyle = 3
    DrawPicture.Line (10, yHeight)-((xWidth + 10), yHeight)
    DrawPicture.DrawStyle = 0
    DrawPicture.Line (xWidth, 10)-(xWidth, (yHeight + 10))
    Dim i As Integer
    For i = 0 To 4
        DrawPicture.Line (xWidth, 10)-(xWidth - i, 20)
        DrawPicture.Line (xWidth, 10)-(xWidth + i, 20)
        DrawPicture.Line (xWidth + 33, yHeight)-(xWidth + 23, yHeight - i)
        DrawPicture.Line (xWidth + 33, yHeight)-(xWidth + 23, yHeight + i)
    Next
    DrawPicture.Line (xWidth + 11, yHeight)-(xWidth + 14, yHeight)
    DrawPicture.Line (xWidth + 17, yHeight)-(xWidth + 25, yHeight)
    DrawPicture.DrawWidth = 2
    DrawPicture.Line (xWidth + 10, 11)-(xWidth + 16, 19)
    DrawPicture.Line (xWidth + 16, 11)-(xWidth + 10, 19)
    DrawPicture.Line (xWidth + 25, yHeight + 10)-(xWidth + 31, yHeight + 10)
    DrawPicture.Line (xWidth + 31, yHeight + 10)-(xWidth + 25, yHeight + 18)
    DrawPicture.Line (xWidth + 25, yHeight + 18)-(xWidth + 31, yHeight + 18)
    DrawPicture.DrawWidth = 1
    xStartPosition = (xWidth + (250 * dScaleFactor))
    yStartPosition = (yHeight - (125 * dScaleFactor))
    xEndPosition = xWidth
    yEndPosition = yHeight
    intColor = 0
End Sub

Private Sub LineByPixel(DrawPicture As PictureBox, xStartPos As Double, yStartPos As Double, xEndPos As Double, yEndPos As Double, LineColor As Long)
    Dim xStep As Integer
    Dim yStep As Double
    Dim xDiff As Double
    Dim yDiff As Double
    Dim xyDiff As Double
    Dim xStart As Double
    Dim yStart As Double
    Dim i As Long
    xStartPos = Int(xStartPos + 0.5)
    yStartPos = Int(yStartPos + 0.5)
    xEndPos = Int(xEndPos + 0.5)
    yEndPos = Int(yEndPos + 0.5)
    If xStartPos = xEndPos And yStartPos = yEndPos Then
        Exit Sub
    End If
    If Round(xStartPos, 3) <= Round(xEndPos, 3) Then
        xStep = 1
        xDiff = Round(xEndPos, 3) - Round(xStartPos, 3)
        xStart = Round(xStartPos, 3)
    Else
        xStep = -1
        xDiff = Round(xStartPos, 3) - Round(xEndPos, 3)
        xStart = Round(xStartPos, 3)
    End If
    If Round(yStartPos, 3) <= Round(yEndPos, 3) Then
        yStep = 1
        yDiff = Round(yEndPos, 3) - Round(yStartPos, 3)
        yStart = Round(yStartPos, 3)
    Else
        yStep = -1
        yDiff = Round(yStartPos, 3) - Round(yEndPos, 3)
        yStart = Round(yStartPos, 3)
    End If
    On Error GoTo errHandle
    'If xDiff = yDiff Then Debug.Print "gelijk"
    'Debug.Print DrawPicture.DrawStyle
    If xDiff >= yDiff Then
        xyDiff = yDiff / xDiff
        For i = Round(xStartPos, 3) To Round(xEndPos, 3) Step xStep
            'Debug.Print i, yStart + IIf(yStep = "1", IIf(xStep = "1", ((i - xStart) * xyDiff), ((xStart - i) * xyDiff)), IIf(xStep = "1", ((xStart - i) * xyDiff), ((i - xStart) * xyDiff)))
            SetPixel DrawPicture.hdc, i, Int(yStart + 0.5 + IIf(yStep = "1", IIf(xStep = "1", ((i - xStart) * xyDiff), ((xStart - i) * xyDiff)), IIf(xStep = "1", ((xStart - i) * xyDiff), ((i - xStart) * xyDiff)))), lLineColor
            DrawPicture.Refresh
            Wait lTimeMilliSeconds
        Next i
    Else
        xyDiff = xDiff / yDiff
        For i = Round(yStartPos, 3) To Round(yEndPos, 3) Step yStep
            'Debug.Print xStart + IIf(xStep = "1", IIf(yStep = "1", ((i - yStart) * xyDiff), ((yStart - i) * xyDiff)), IIf(yStep = "1", ((yStart - i) * xyDiff), ((i - yStart) * xyDiff))), i
            SetPixel DrawPicture.hdc, Int(xStart + 0.5 + IIf(xStep = "1", IIf(yStep = "1", ((i - yStart) * xyDiff), ((yStart - i) * xyDiff)), IIf(yStep = "1", ((yStart - i) * xyDiff), ((i - yStart) * xyDiff)))), i, lLineColor
            DrawPicture.Refresh
            Wait lTimeMilliSeconds
        Next i
    End If
    Exit Sub
errHandle:
    xyDiff = 1
    Debug.Print "Error into LineByPixel", frmGraphic.Caption
'    Resume Next
End Sub
