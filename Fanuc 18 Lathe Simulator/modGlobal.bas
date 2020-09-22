Attribute VB_Name = "modGlobal"
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public iOpenFiles As Integer
Public iRepeated() As Integer
Public iLastReaded() As Integer

Public intColor  As Integer
Public lLineColor As Long
Public lTimeMilliSeconds As Long

Public dScaleFactor As Double

Public xWidth As Integer 'xNullPoint PictureBox -
Public yHeight As Integer 'yNullPoint Picturebox |

'Start - End points for DrawContour
Public xStartPosition As Double
Public yStartPosition As Double
Public xEndPosition As Double
Public yEndPosition As Double

Public Sub Wait(TimeMilliSeconds As Long)
    If TimeMilliSeconds <> 0 Then
        Dim EndTime As Long
        EndTime = GetTickCount + TimeMilliSeconds
        Do Until GetTickCount > EndTime
            DoEvents
        Loop
    End If
End Sub

Public Function CheckCommand(sData As String) As String
    Dim sData1 As String
    Dim sChar As String
    Dim iCount As Integer
    sData = UCase(sData)
    sData1 = ""
    If InStr(1, sData, ":0") <> 0 Then GoTo endFunction
    For iCount = 1 To Len(sData)
        sChar = Mid(sData, iCount, 1)
        If Not IsNumeric(sChar) And iCount <> 1 And sChar <> "-" Then
            If InStr(1, sData, "(") = 0 Then
                If sChar <> "." Then
                    If sChar <> " " Then
                        sData1 = sData1 & " " & sChar
                    End If
                Else
                    If IsNumeric(Mid(sData, iCount + 1, 1)) Then
                        sData1 = sData1 & sChar
                    End If
                End If
            Else
                sData1 = sData1 & sChar
            End If
        Else
            sData1 = sData1 & sChar
        End If
    Next
    If Mid(sData1, Len(sData1), 1) <> ";" And Left((sData1), 1) <> "%" Then
        If Mid(sData1, Len(sData1), 1) <> " " Then
            sData1 = sData1 & " ;"
        Else
            sData1 = sData1 & ";"
        End If
    End If
    sData = sData1
    Dim iGPos As Integer
    If InStr(1, sData, "G") <> 0 Then
        iGPos = InStr(1, sData, "G")
        If IsNumeric(Mid(sData, iGPos + 1, 1)) Then
            If Mid(sData, iGPos + 2, 1) = " " Then
                sData = Left(sData, iGPos) & "0" & Mid(sData, iGPos + 1)
            End If
        End If
    End If
    If InStr(1, sData, "M") <> 0 Then
        If InStr(1, sData, "M98") <> 0 Then
            iGPos = InStr(1, sData, "P") + 1
            For iCount = iGPos To iGPos + 4
                If Mid(sData, iCount, 1) = " " Then
                    sData = Left(sData, iCount - 1) & Mid(sData, iCount + 1)
                End If
            Next
            GoTo endFunction
        End If
        If InStr(1, sData, "M99") <> 0 Then GoTo endFunction
        If InStr(1, sData, "M30") <> 0 Then GoTo endFunction
        iGPos = InStr(1, sData, "M")
        If IsNumeric(Mid(sData, iGPos + 1, 1)) Then
            If Mid(sData, iGPos + 2, 1) = " " Then
                sData = Left(sData, iGPos) & "0" & Mid(sData, iGPos + 1)
            End If
        End If
    End If
endFunction:
    CheckCommand = sData
End Function

Public Function FileExist(Bestand$)
    If InStr(1, Bestand$, ".") = 0 Then FileExist = False: Exit Function:
    On Error Resume Next
    Call FileLen(Bestand$)
    FileExist = (Err = 0)
End Function

Public Function GetCoordinates(sData As String, iStart As Integer, Optional blnToScale As Boolean = True) As Double
    Dim sCoordinate As String
    Dim iLength As Integer
    On Error GoTo errHandle
    iLength = iStart + 1
    Do
        sCoordinate = sCoordinate & Mid(sData, iLength, 1)
        iLength = iLength + 1
    Loop Until Mid(sData, iLength, 1) = " " Or iLength >= Len(sData)
    If InStr(1, sCoordinate, ".") <> 0 Then
        iLength = InStr(1, sCoordinate, ".")
        sCoordinate = Left(sCoordinate, iLength - 1) & "," & Mid(sCoordinate, iLength + 1)
    End If
    If blnToScale = True Then
        GetCoordinates = sCoordinate * dScaleFactor
    Else
        GetCoordinates = sCoordinate
    End If
    Exit Function
errHandle:
    GetCoordinates = 0
End Function

Public Function GetPath(ByVal FileName As String) As String
    Dim p As Integer, j As Integer
    p = 0
    For j = Len(FileName) To 1 Step -1
        If Mid(FileName, j, 1) = "\" Then p = j: Exit For
    Next
    If p > 0 Then GetPath = Left$(FileName, p) Else GetPath = ""
End Function

Public Function LenCoordinates(sData As String, iStart As Integer, Optional blnToScale As Boolean = True)
    Dim sCoordinate As String
    Dim iLength As Integer
    On Error GoTo errHandle
    iLength = iStart + 1
    Do
        sCoordinate = sCoordinate & Mid(sData, iLength, 1)
        iLength = iLength + 1
    Loop Until Mid(sData, iLength, 1) = " " Or iLength >= Len(sData)
    If blnToScale = True Then
        sCoordinate = sCoordinate * dScaleFactor
    Else
        sCoordinate = sCoordinate
    End If
    LenCoordinates = Len(sCoordinate)
    Exit Function
errHandle:
    LenCoordinates = ""
End Function

Public Sub SetStartPoint(sData As String)
    If InStr(1, sData, "X") <> 0 And InStr(1, sData, "Z") <> 0 Then
        xEndPosition = xWidth + GetCoordinates(sData, InStr(1, sData, "Z"))
        yEndPosition = yHeight - (GetCoordinates(sData, InStr(1, sData, "X")) / 2)
    ElseIf InStr(1, sData, "Z") <> 0 Then
        xEndPosition = xWidth + GetCoordinates(sData, InStr(1, sData, "Z"))
    Else
        yEndPosition = yHeight - (GetCoordinates(sData, InStr(1, sData, "X")) / 2)
    End If
    xStartPosition = xEndPosition
    yStartPosition = yEndPosition
End Sub
