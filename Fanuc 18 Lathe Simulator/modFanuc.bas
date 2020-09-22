Attribute VB_Name = "modFanuc"
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

'Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Type Rect
    Left As Double
    Top As Double
    Right As Double
    Bottom As Double
End Type

Public Sub DoCommandFanuc18TC(mForm As Form, mListbox As ListBox, mListIndex As Integer, mResultPicture As PictureBox, mCalcPicture As PictureBox)
    'Debug.Print mForm.Name
    'Debug.Print mListbox.Name
    'Debug.Print mListIndex
    'Debug.Print mResultPicture.Name
    'Debug.Print mCalcPicture.Name
    Dim sData As String
    Dim sCommand As String
    sData = mListbox.List(mListIndex)
    If InStr(1, sData, "G") <> 0 Then
        Dim iPosition As Integer
        iPosition = InStr(1, sData, "G")
        sCommand = Mid(sData, iPosition, 3)
        If IsNumeric(Right(sCommand, 2)) Then
            'start position
            Dim dsXpos As Double
            Dim dsYpos As Double
            'end position
            Dim dEXpos As Double
            Dim dEYpos As Double
            'automatic pass and back value
            Dim dPassValueX As Double
            Dim dPassValueZ As Double
            Dim dBackValue As Double
            'feed
            Dim dFeed As Double
            'extra on X and Y
            Dim dExtraX As Double
            Dim dExtraY As Double
            'inner or outer
            Dim blnInner As Boolean
            'automatic start and end line
            Dim SLine As String
            Dim ELine As String
            Dim sLockPoint As String
            'For G70 G71 G72
            Dim CuttingRegio As Rect
            Dim pStartX As Double
            Dim pStartY As Double
            Dim pEndX As Double
            Dim pEndY As Double
            Dim xCor As Long
            Dim yCor As Long
            'End G70 G71 G72
            Dim iCount As Integer
            Dim oldListIndex As Integer
            Dim blnAddTo As Boolean
            Select Case Mid(sCommand, 2, 2)
                Case "00"
                Case "02", "03"
                Case "04"
                Case "09"
                Case "28"
                Case "32"
                Case "50"
                Case "68"
                Case "69"
                Case "70"
                    If InStr(1, sData, ":0") <> 0 Then GoTo EndSelectCase
                    dsXpos = GetCoordinates(mListbox.List(mListIndex - 1), InStr(1, mListbox.List(mListIndex - 1), "Z"), False)
                    dsYpos = GetCoordinates(mListbox.List(mListIndex - 1), InStr(1, mListbox.List(mListIndex - 1), "X"), False)
                    SLine = "N" & GetCoordinates(sData, InStr(1, sData, "P"), False)
                    ELine = "N" & GetCoordinates(sData, InStr(1, sData, "Q"), False)
                    sData = "G00 X" & dsYpos & " Z" & dsXpos & " ;"
                    PaintResult sData, mForm, mListbox, mResultPicture, False
                    iCount = 0
                    oldListIndex = mListbox.ListIndex
                    Do
                        sData = mListbox.List(iCount)
                        If InStr(1, sData, SLine) <> 0 Then
                            blnAddTo = True
                        End If
                        If blnAddTo = True Then
                            mListbox.ListIndex = iCount
                            mListbox.TopIndex = mListbox.ListIndex
                            frmGraphic.StopByLine sData
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                        End If
                        iCount = iCount + 1
                    Loop Until InStr(1, sData, ELine) <> 0
                    blnAddTo = False
                    sData = "G00 X" & dsYpos & " Z" & dsXpos & " ;"
                    mListbox.ListIndex = oldListIndex
                    mListbox.TopIndex = mListbox.ListIndex
                Case "71"
                    If InStr(1, sData, ":0") <> 0 Then GoTo EndSelectCase
                    If InStr(1, sData, "R") <> 0 Then
                        dPassValueX = GetCoordinates(sData, InStr(1, sData, "U"), True)
                        dBackValue = GetCoordinates(sData, InStr(1, sData, "R"), False)
                        dsXpos = GetCoordinates(mListbox.List(mListIndex - 1), InStr(1, mListbox.List(mListIndex - 1), "Z"), False)
                        dsYpos = GetCoordinates(mListbox.List(mListIndex - 1), InStr(1, mListbox.List(mListIndex - 1), "X"), False)
                        mListbox.ListIndex = mListbox.ListIndex + 1
                        sData = mListbox.List(mListbox.ListIndex)
                        mListbox.TopIndex = mListbox.ListIndex
                        frmGraphic.StopByLine sData
                        SLine = "N" & GetCoordinates(sData, InStr(1, sData, "P"), False)
                        ELine = "N" & GetCoordinates(sData, InStr(1, sData, "Q"), False)
                        dExtraX = GetCoordinates(sData, InStr(1, sData, "W"), False)
                        dExtraY = GetCoordinates(sData, InStr(1, sData, "U"), False)
                        blnInner = IIf(InStr(1, sData, "-") <> 0, True, False)
                        dFeed = GetCoordinates(sData, InStr(1, sData, "F"), False)
                        sLockPoint = "G00 X" & dsYpos & " Z" & dsXpos '& " F" & Trim(dFeed) & " ;"
                        SetStartPoint sLockPoint
                        CuttingRegio.Top = yStartPosition '- 1
                        CuttingRegio.Right = xStartPosition '+ 1
                        mForm!lstAutomatic.Clear
                        mListbox.ListIndex = mListbox.ListIndex + 1
                        sData = AddForCut(mListbox.List(mListbox.ListIndex), 2, 0, 0, dExtraY, dExtraX, 2)
                        'Debug.Print sData
                        mListbox.TopIndex = mListbox.ListIndex
                        mForm!lstAutomatic.AddItem sData
                        frmGraphic.StopByLine sData
                        SetStartPoint sData
                        CuttingRegio.Bottom = yStartPosition '+ 1
                        Do
                            mListbox.ListIndex = mListbox.ListIndex + 1
                            sData = AddForCut(mListbox.List(mListbox.ListIndex), 2, 0, 0, dExtraY, dExtraX, 2)
                            mListbox.TopIndex = mListbox.ListIndex
                            frmGraphic.StopByLine sData
                            'Debug.Print sData
                            mForm!lstAutomatic.AddItem sData
                        Loop Until InStr(1, sData, ELine) <> 0
                        mCalcPicture.Cls
                        mCalcPicture.ForeColor = vbBlack
                        pStartX = xStartPosition
                        pStartY = yStartPosition
                        pEndX = xEndPosition
                        pEndY = yEndPosition
                        PaintForCalc mForm!lstAutomatic, mCalcPicture, True
                        xStartPosition = pStartX
                        yStartPosition = pStartY
                        xEndPosition = pEndX
                        yEndPosition = pEndY
                        PaintForCalc mForm!lstAutomatic, mCalcPicture, False
                        CuttingRegio.Left = xEndPosition '- 1
                        SetStartPoint sLockPoint
                        If CuttingRegio.Top < CuttingRegio.Bottom Then
                            'Rectangle mCalcPicture.hdc, CuttingRegio.Left - 1, CuttingRegio.Top - 1, CuttingRegio.Right + 1, CuttingRegio.Bottom + 2
                            For yCor = CuttingRegio.Top + dPassValueX To CuttingRegio.Bottom Step dPassValueX
                                For xCor = CuttingRegio.Right - 2 To CuttingRegio.Left Step -1
                                    If GetPixel(mCalcPicture.hdc, xCor, yCor) = vbBlack Then
                                        sData = "G01 X" & Round(((yHeight - yCor) / dScaleFactor) * 2, 2) & " ;"
                                        PaintResult sData, mForm, mListbox, mResultPicture, False
                                        sData = "G01 Z-" & Round(((xWidth - xCor) / dScaleFactor), 2) - dExtraX & " ;"
                                        PaintResult sData, mForm, mListbox, mResultPicture, False
                                        sData = "G00 U" & dBackValue * 2 & " W" & dBackValue & " ;"
                                        PaintResult sData, mForm, mListbox, mResultPicture, False
                                        sData = "G00 Z" & dsXpos & " ;"
                                        PaintResult sData, mForm, mListbox, mResultPicture, False
                                        Exit For
                                    End If
                                Next
                            Next
                        Else
                            'Rectangle mCalcPicture.hdc, CuttingRegio.Left - 1, CuttingRegio.Top + 1, CuttingRegio.Right + 1, CuttingRegio.Bottom + 1
                            For yCor = CuttingRegio.Top - dPassValueX To CuttingRegio.Bottom Step -dPassValueX
                                For xCor = CuttingRegio.Right - 2 To CuttingRegio.Left Step -1
                                    If GetPixel(mCalcPicture.hdc, xCor, yCor) = vbBlack Then
                                        sData = "G01 X" & Round(((yHeight - yCor) / dScaleFactor) * 2, 2) & " ;"
                                        'Debug.Print sData
                                        PaintResult sData, mForm, mListbox, mResultPicture, False
                                        sData = "G01 Z-" & Round(((xWidth - xCor) / dScaleFactor), 2) - dExtraX & " ;"
                                        'Debug.Print sData
                                        PaintResult sData, mForm, mListbox, mResultPicture, False
                                        sData = "G00 U-" & dBackValue * 2 & " W" & dBackValue & " ;"
                                        'Debug.Print sData
                                        PaintResult sData, mForm, mListbox, mResultPicture, False
                                        sData = "G00 Z" & dsXpos & " ;"
                                        'Debug.Print sData
                                        PaintResult sData, mForm, mListbox, mResultPicture, False
                                        Exit For
                                    End If
                                Next
                            Next
                        End If
                        PaintResult sLockPoint, mForm, mListbox, mResultPicture, False
                        iCount = 0
                        Do
                            sData = mListbox.List(iCount)
                            If InStr(1, sData, SLine) <> 0 Then
                                blnAddTo = True
                            End If
                            If blnAddTo = True Then
                                mListbox.ListIndex = iCount
                                mListbox.TopIndex = mListbox.ListIndex
                                frmGraphic.StopByLine sData
                                sData = AddForCut(sData, 2, 0, 0, dExtraY, dExtraX, 2)
                                PaintResult sData, mForm, mListbox, mResultPicture, False
                            End If
                            iCount = iCount + 1
                        Loop Until InStr(1, sData, ELine) <> 0
                        blnAddTo = False
                        sData = sLockPoint
                        iLastReaded(iOpenFiles) = mListbox.ListIndex
                    End If
                Case "72"
                    If InStr(1, sData, ":0") <> 0 Then GoTo EndSelectCase
                    If InStr(1, sData, "R") <> 0 Then
                        dPassValueZ = GetCoordinates(sData, InStr(1, sData, "W"), True)
                        dBackValue = GetCoordinates(sData, InStr(1, sData, "R"), False)
                        dsXpos = GetCoordinates(mListbox.List(mListIndex - 1), InStr(1, mListbox.List(mListIndex - 1), "Z"), False)
                        dsYpos = GetCoordinates(mListbox.List(mListIndex - 1), InStr(1, mListbox.List(mListIndex - 1), "X"), False)
                        mListbox.ListIndex = mListbox.ListIndex + 1
                        sData = mListbox.List(mListbox.ListIndex)
                        mListbox.TopIndex = mListbox.ListIndex
                        frmGraphic.StopByLine sData
                        SLine = "N" & GetCoordinates(sData, InStr(1, sData, "P"), False)
                        ELine = "N" & GetCoordinates(sData, InStr(1, sData, "Q"), False)
                        dExtraX = GetCoordinates(sData, InStr(1, sData, "W"), False)
                        dExtraY = GetCoordinates(sData, InStr(1, sData, "U"), False)
                        blnInner = IIf(InStr(1, sData, "-") <> 0, True, False)
                        dFeed = GetCoordinates(sData, InStr(1, sData, "F"), False)
                        sLockPoint = "G00 X" & dsYpos & " Z" & dsXpos '& " F" & Trim(dFeed) & " ;"
                        SetStartPoint sLockPoint
                        CuttingRegio.Top = yStartPosition '- 1
                        CuttingRegio.Right = xStartPosition '+ 1
                        mForm!lstAutomatic.Clear
                        mListbox.ListIndex = mListbox.ListIndex + 1
                        sData = AddForCut(mListbox.List(mListbox.ListIndex), 2, 0, 0, dExtraY, dExtraX, 2)
                        'Debug.Print sData
                        mListbox.TopIndex = mListbox.ListIndex
                        mForm!lstAutomatic.AddItem sData
                        frmGraphic.StopByLine sData
                        SetStartPoint sData
                        'mForm!lstAutomatic.Clear
                        'mListbox.ListIndex = mListbox.ListIndex + 1
                        'sData = mListbox.List(mListbox.ListIndex)
                        'mListbox.TopIndex = mListbox.ListIndex
                        'frmGraphic.StopByLine sData
                        'SetStartPoint sData
                        CuttingRegio.Bottom = yStartPosition '+ 1
                        Do
                            mListbox.ListIndex = mListbox.ListIndex + 1
                            sData = AddForCut(mListbox.List(mListbox.ListIndex), 2, 0, 0, dExtraY, dExtraX, 2)
                            mListbox.TopIndex = mListbox.ListIndex
                            frmGraphic.StopByLine sData
                            mForm!lstAutomatic.AddItem sData
                        Loop Until InStr(1, sData, ELine) <> 0
                        mCalcPicture.Cls
                        mCalcPicture.ForeColor = vbBlack
                        pStartX = xStartPosition
                        pStartY = yStartPosition
                        pEndX = xEndPosition
                        pEndY = yEndPosition
                        PaintForCalc mForm!lstAutomatic, mCalcPicture, True
                        xStartPosition = pStartX
                        yStartPosition = pStartY
                        xEndPosition = pEndX
                        yEndPosition = pEndY
                        PaintForCalc mForm!lstAutomatic, mCalcPicture, False
                        CuttingRegio.Left = xEndPosition '- 1
                        SetStartPoint sLockPoint
                        If CuttingRegio.Top < CuttingRegio.Bottom Then
                            'Rectangle mCalcPicture.hdc, CuttingRegio.Left - 1, CuttingRegio.Top - 1, CuttingRegio.Right + 1, CuttingRegio.Bottom + 2
                            For xCor = CuttingRegio.Right - dPassValueZ To CuttingRegio.Left Step -dPassValueZ
                                For yCor = CuttingRegio.Top + 2 To CuttingRegio.Bottom Step 1
                                    If GetPixel(mCalcPicture.hdc, xCor, yCor) = vbBlack Then
                                        'SetPixel mCalcPicture.hdc, xCor, yCor, vbRed
                                        'mCalcPicture.Refresh
                                        sData = "G01 Z-" & Round(((xWidth - xCor) / dScaleFactor), 2) - dExtraX & " ;"
                                        PaintResult sData, mForm, mListbox, mResultPicture, False
                                        sData = "G01 X" & Round(((yHeight - yCor) / dScaleFactor) * 2, 2) & " ;"
                                        PaintResult sData, mForm, mListbox, mResultPicture, False
                                        sData = "G00 U" & dBackValue * 2 & " W" & dBackValue & " ;"
                                        PaintResult sData, mForm, mListbox, mResultPicture, False
                                        sData = "G00 X" & dsYpos & " ;"
                                        PaintResult sData, mForm, mListbox, mResultPicture, False
                                        Exit For
                                    'Else
                                        'SetPixel mCalcPicture.hdc, xCor, yCor, vbYellow
                                        'mCalcPicture.Refresh
                                    End If
                                Next
                            Next
                        Else
                            'Rectangle mCalcPicture.hdc, CuttingRegio.Left - 1, CuttingRegio.Top + 1, CuttingRegio.Right + 1, CuttingRegio.Bottom + 1
                            For xCor = CuttingRegio.Right - dPassValueZ To CuttingRegio.Left Step -dPassValueZ
                                For yCor = CuttingRegio.Top - 2 To CuttingRegio.Bottom Step -1
                                    If GetPixel(mCalcPicture.hdc, xCor, yCor) = vbBlack Then
                                        'SetPixel mCalcPicture.hdc, xCor, yCor, vbRed
                                        'mCalcPicture.Refresh
                                        sData = "G01 Z-" & Round(((xWidth - xCor) / dScaleFactor), 2) - dExtraX & " ;"
                                        'Debug.Print sData
                                        PaintResult sData, mForm, mListbox, mResultPicture, False
                                        sData = "G01 X" & Round(((yHeight - yCor) / dScaleFactor) * 2, 2) & " ;"
                                        'Debug.Print sData
                                        PaintResult sData, mForm, mListbox, mResultPicture, False
                                        sData = "G00 U-" & dBackValue * 2 & " W" & dBackValue & " ;"
                                        'Debug.Print sData
                                        PaintResult sData, mForm, mListbox, mResultPicture, False
                                        sData = "G00 X" & dsYpos & " ;"
                                        'Debug.Print sData
                                        PaintResult sData, mForm, mListbox, mResultPicture, False
                                        Exit For
                                    'Else
                                        'SetPixel mCalcPicture.hdc, xCor, yCor, vbYellow
                                        'mCalcPicture.Refresh
                                    End If
                                Next
                            Next
                        End If
                        PaintResult sLockPoint, mForm, mListbox, mResultPicture, False
                        iCount = 0
                        Do
                            sData = mListbox.List(iCount)
                            If InStr(1, sData, SLine) <> 0 Then
                                blnAddTo = True
                            End If
                            If blnAddTo = True Then
                                mListbox.ListIndex = iCount
                                mListbox.TopIndex = mListbox.ListIndex
                                frmGraphic.StopByLine sData
                                sData = AddForCut(sData, 2, 0, 0, dExtraY, dExtraX, 2)
                                PaintResult sData, mForm, mListbox, mResultPicture, False
                            End If
                            iCount = iCount + 1
                        Loop Until InStr(1, sData, ELine) <> 0
                        blnAddTo = False
                        sData = sLockPoint
                        iLastReaded(iOpenFiles) = mListbox.ListIndex
                    End If
                Case "73"
                    If InStr(1, sData, ":0") <> 0 Then GoTo EndSelectCase
                    dsXpos = GetCoordinates(mListbox.List(mListIndex - 1), InStr(1, mListbox.List(mListIndex - 1), "Z"), False)
                    dsYpos = GetCoordinates(mListbox.List(mListIndex - 1), InStr(1, mListbox.List(mListIndex - 1), "X"), False)
                    sLockPoint = "G00 X" & dsYpos & " Z" & dsXpos & " ;"
                    dPassValueX = GetCoordinates(sData, InStr(1, sData, "U"), False)
                    dPassValueZ = GetCoordinates(sData, InStr(1, sData, "W"), False)
                    dBackValue = GetCoordinates(sData, InStr(1, sData, "R"), False)
                    mListbox.ListIndex = mListbox.ListIndex + 1
                    sData = mListbox.List(mListbox.ListIndex)
                    mListbox.TopIndex = mListbox.ListIndex
                    frmGraphic.StopByLine sData
                    SLine = "N" & GetCoordinates(sData, InStr(1, sData, "P"), False)
                    ELine = "N" & GetCoordinates(sData, InStr(1, sData, "Q"), False)
                    dExtraX = GetCoordinates(sData, InStr(1, sData, "W"), False)
                    dExtraY = GetCoordinates(sData, InStr(1, sData, "U"), False)
                    dFeed = GetCoordinates(sData, InStr(1, sData, "F"), False)
                    iCount = 0
                    oldListIndex = mListbox.ListIndex
                    Dim iCut As Integer
                    For iCut = 1 To dBackValue
                        iCount = 0
                        Do
                            sData = mListbox.List(iCount)
                            If InStr(1, sData, SLine) <> 0 Then
                                blnAddTo = True
                            End If
                            If blnAddTo = True Then
                                mListbox.ListIndex = iCount
                                mListbox.TopIndex = mListbox.ListIndex
                                frmGraphic.StopByLine sData
                                sData = AddForCut(sData, dBackValue, dPassValueX, dPassValueZ, dExtraY, dExtraX, iCut)
                                PaintResult sData, mForm, mListbox, mResultPicture, False
                            End If
                            iCount = iCount + 1
                        Loop Until InStr(1, sData, ELine) <> 0
                        blnAddTo = False
                        PaintResult sLockPoint, mForm, mListbox, mResultPicture, False
                    Next
                    sData = "G00 X" & dsYpos & " Z" & dsXpos & " ;"
                    iLastReaded(iOpenFiles) = mListbox.ListIndex
                Case "74"
                Case "75"
                Case "76"
                    If InStr(1, sData, ":0") <> 0 Then GoTo EndSelectCase
                    If InStr(1, sData, "P") <> 0 Then
                        If InStr(1, mListbox.List(mListIndex - 1), "G97") <> 0 Then
                            dsYpos = GetCoordinates(mListbox.List(mListIndex - 2), InStr(1, mListbox.List(mListIndex - 2), "X"), False)
                            dsXpos = GetCoordinates(mListbox.List(mListIndex - 2), InStr(1, mListbox.List(mListIndex - 2), "Z"), False)
                        Else
                            dsYpos = GetCoordinates(mListbox.List(mListIndex - 1), InStr(1, mListbox.List(mListIndex - 1), "X"), False)
                            dsXpos = GetCoordinates(mListbox.List(mListIndex - 1), InStr(1, mListbox.List(mListIndex - 1), "Z"), False)
                        End If
                        sLockPoint = "G00 X" & dsYpos & " Z" & dsXpos & " ;"
                        dBackValue = GetCoordinates(sData, InStr(1, sData, "P"), False)
                        dPassValueX = GetCoordinates(sData, InStr(1, sData, "Q"), False)
                        dExtraX = GetCoordinates(sData, InStr(1, sData, "R"), False)
                        mListbox.ListIndex = mListbox.ListIndex + 1
                        sData = mListbox.List(mListbox.ListIndex)
                        mListbox.TopIndex = mListbox.ListIndex
                        frmGraphic.StopByLine sData
                        dEYpos = GetCoordinates(sData, InStr(1, sData, "X"), False)
                        dEXpos = GetCoordinates(sData, InStr(1, sData, "Z"), False)
                        dExtraY = GetCoordinates(sData, InStr(1, sData, "R"), False)
                        blnInner = IIf(InStr(1, sData, "-") <> 0, True, False)
                        dFeed = GetCoordinates(sData, InStr(1, sData, "F"), False)
                        If dsYpos < dEYpos Then
                            pStartX = dEYpos - (((GetCoordinates(sData, InStr(1, sData, "P"), False)) * 2) / 1000)
                            pStartX = pStartX + (((GetCoordinates(sData, InStr(1, sData, "Q"), False)) * 2) / 1000)
                            sData = "G00 X" & pStartX & " ;"
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                            sData = "G01 X" & (pStartX + GetCoordinates(sData, InStr(1, sData, "R"), False)) & " Z" & (dEXpos + (Left(Right(dBackValue, 4), 2) / 10 * GetCoordinates(sData, InStr(1, sData, "F"), False))) & " ;"
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                            sData = "G00 X" & dsYpos & " ;"
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                            sData = "G00 Z" & dsXpos & " ;"
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                            pStartX = pStartX + ((dPassValueX * 2) / 1000)
                            For yCor = pStartX * 1000 To dEYpos * 1000 Step dPassValueX * 2
                                sData = "G00 X" & yCor / 1000 & " ;"
                                PaintResult sData, mForm, mListbox, mResultPicture, False
                                sData = "G01 X" & (yCor / 1000 + GetCoordinates(sData, InStr(1, sData, "R"), False)) & " Z" & (dEXpos + (Left(Right(dBackValue, 4), 2) / 10 * GetCoordinates(sData, InStr(1, sData, "F"), False))) & " ;"
                                PaintResult sData, mForm, mListbox, mResultPicture, False
                                sData = "G00 X" & dsYpos & " ;"
                                PaintResult sData, mForm, mListbox, mResultPicture, False
                                sData = "G00 Z" & dsXpos & " ;"
                                PaintResult sData, mForm, mListbox, mResultPicture, False
                            Next
                        Else
                            pStartX = dEYpos + (((GetCoordinates(sData, InStr(1, sData, "P"), False)) * 2) / 1000)
                            pStartX = pStartX - (((GetCoordinates(sData, InStr(1, sData, "Q"), False)) * 2) / 1000)
                            sData = "G00 X" & pStartX & " ;"
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                            sData = "G01 X" & (pStartX + GetCoordinates(sData, InStr(1, sData, "R"), False)) & " Z" & (dEXpos + (Left(Right(dBackValue, 4), 2) / 10 * GetCoordinates(sData, InStr(1, sData, "F"), False))) & " ;"
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                            sData = "G00 X" & dsYpos & " ;"
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                            sData = "G00 Z" & dsXpos & " ;"
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                            pStartX = pStartX - ((dPassValueX * 2) / 1000)
                            For yCor = pStartX * 1000 To dEYpos * 1000 Step -dPassValueX * 2
                                sData = "G00 X" & yCor / 1000 & " ;"
                                PaintResult sData, mForm, mListbox, mResultPicture, False
                                sData = "G01 X" & (yCor / 1000 + GetCoordinates(sData, InStr(1, sData, "R"), False)) & " Z" & (dEXpos + (Left(Right(dBackValue, 4), 2) / 10 * GetCoordinates(sData, InStr(1, sData, "F"), False))) & " ;"
                                PaintResult sData, mForm, mListbox, mResultPicture, False
                                sData = "G00 X" & dsYpos & " ;"
                                PaintResult sData, mForm, mListbox, mResultPicture, False
                                sData = "G00 Z" & dsXpos & " ;"
                                PaintResult sData, mForm, mListbox, mResultPicture, False
                            Next
                        End If
                        sData = sLockPoint
                        iLastReaded(iOpenFiles) = mListbox.ListIndex
                    End If
                Case "80"
                Case "83"
                Case "84"
                Case "85"
                Case "87"
                Case "88"
                Case "89"
                Case "90", "92"
                    If InStr(1, sData, ":0") <> 0 Then GoTo EndSelectCase
                    dFeed = GetCoordinates(sData, InStr(1, sData, "F"), False)
                    dsYpos = GetCoordinates(mListbox.List(mListIndex - 1), InStr(1, mListbox.List(mListIndex - 1), "X"), False)
                    dsXpos = GetCoordinates(mListbox.List(mListIndex - 1), InStr(1, mListbox.List(mListIndex - 1), "Z"), False)
                    dExtraX = 0
                    If InStr(1, sData, "R") <> 0 Then
                        dExtraX = GetCoordinates(sData, InStr(1, sData, "R"), False) * 2
                    End If
                    sLockPoint = "G01 Z" & GetCoordinates(sData, InStr(1, sData, "Z"), False)
                    sData = "G00 X" & GetCoordinates(sData, InStr(1, sData, "X"), False) + dExtraX & " ;"
                    PaintResult sData, mForm, mListbox, mResultPicture, False
                    sData = sLockPoint & " X" & GetCoordinates(sData, InStr(1, sData, "X"), False) - dExtraX & " F" & Trim(dFeed) & " ;"
                    PaintResult sData, mForm, mListbox, mResultPicture, False
                    sData = "X" & dsYpos & " ;"
                    PaintResult sData, mForm, mListbox, mResultPicture, False
                    sData = "G00 X" & dsYpos & " Z" & dsXpos & " ;"
                    PaintResult sData, mForm, mListbox, mResultPicture, False
                    Do
                        mListbox.ListIndex = mListbox.ListIndex + 1
                        mListbox.TopIndex = mListbox.ListIndex
                        sData = mListbox.List(mListbox.ListIndex)
                        frmGraphic.StopByLine sData
                        If InStr(1, sData, "G") = 0 Then
                            sData = "X" & GetCoordinates(sData, InStr(1, sData, "X"), False) + dExtraX & " ;"
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                            sData = sLockPoint & " X" & GetCoordinates(sData, InStr(1, sData, "X"), False) - dExtraX & " ;"
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                            sData = "X" & dsYpos & " ;"
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                            sData = "G00 Z" & dsXpos & " ;"
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                            sData = mListbox.List(mListbox.ListIndex)
                        End If
                    Loop Until InStr(1, sData, "G") <> 0
                    iLastReaded(iOpenFiles) = mListbox.ListIndex
'                Case "92"
'                    If InStr(1, sData, ":0") <> 0 Then GoTo EndSelectCase
'                    dFeed = GetCoordinates(sData, InStr(1, sData, "F"), False)
'                    dSYpos = GetCoordinates(mListbox.List(mListIndex - 1), InStr(1, mListbox.List(mListIndex - 1), "X"), False)
'                    dSXpos = GetCoordinates(mListbox.List(mListIndex - 1), InStr(1, mListbox.List(mListIndex - 1), "Z"), False)
'                    sLockPoint = "G01 Z" & GetCoordinates(sData, InStr(1, sData, "Z"), False) & " F" & Trim(dFeed)
'                    sData = "G00 X" & GetCoordinates(sData, InStr(1, sData, "X"), False) & " ;"
'                    PaintResult sData, mForm, mListbox, mResultPicture, False
'                    sData = sLockPoint & " ;"
'                    PaintResult sData, mForm, mListbox, mResultPicture, False
'                    sData = "X" & dSYpos & " ;"
'                    PaintResult sData, mForm, mListbox, mResultPicture, False
'                    sData = "G00 Z" & dSXpos & " ;"
'                    PaintResult sData, mForm, mListbox, mResultPicture, False
'                    Do
'                        mListbox.ListIndex = mListbox.ListIndex + 1
'                        mListbox.TopIndex = mListbox.ListIndex
'                        sData = mListbox.List(mListbox.ListIndex)
'                        frmGraphic.StopByLine sData
'                        If InStr(1, sData, "G") = 0 Then
'                            sData = "X" & GetCoordinates(sData, InStr(1, sData, "X"), False) & " ;"
'                            PaintResult sData, mForm, mListbox, mResultPicture, False
'                            sData = sLockPoint & " ;"
'                            PaintResult sData, mForm, mListbox, mResultPicture, False
'                            sData = "X" & dSYpos & " ;"
'                            PaintResult sData, mForm, mListbox, mResultPicture, False
'                            sData = "G00 Z" & dSXpos & " ;"
'                            PaintResult sData, mForm, mListbox, mResultPicture, False
'                            sData = mListbox.List(mListbox.ListIndex)
'                        End If
'                    Loop Until InStr(1, sData, "G") <> 0
'                    iLastReaded(iOpenFiles) = mListbox.ListIndex
                Case "94"
                    If InStr(1, sData, ":0") <> 0 Then GoTo EndSelectCase
                    dFeed = GetCoordinates(sData, InStr(1, sData, "F"), False)
                    dsYpos = GetCoordinates(mListbox.List(mListIndex - 1), InStr(1, mListbox.List(mListIndex - 1), "X"), False)
                    dsXpos = GetCoordinates(mListbox.List(mListIndex - 1), InStr(1, mListbox.List(mListIndex - 1), "Z"), False)
                    dExtraX = 0
                    If InStr(1, sData, "R") <> 0 Then
                        dExtraX = GetCoordinates(sData, InStr(1, sData, "R"), False)
                    End If
                    sLockPoint = "G01 X" & GetCoordinates(sData, InStr(1, sData, "X"), False)
                    sData = "G00 Z" & GetCoordinates(sData, InStr(1, sData, "Z"), False) + dExtraX & " ;"
                    PaintResult sData, mForm, mListbox, mResultPicture, False
                    sData = sLockPoint & " Z" & GetCoordinates(sData, InStr(1, sData, "Z"), False) - dExtraX & " F" & dFeed & " ;"
                    PaintResult sData, mForm, mListbox, mResultPicture, False
                    sData = "Z" & dsXpos & " ;"
                    PaintResult sData, mForm, mListbox, mResultPicture, False
                    sData = "G00 X" & dsYpos & " ;"
                    PaintResult sData, mForm, mListbox, mResultPicture, False
                    Do
                        mListbox.ListIndex = mListbox.ListIndex + 1
                        mListbox.TopIndex = mListbox.ListIndex
                        sData = mListbox.List(mListbox.ListIndex)
                        frmGraphic.StopByLine sData
                        If InStr(1, sData, "G") = 0 Then
                            sData = "Z" & GetCoordinates(sData, InStr(1, sData, "Z"), False) + dExtraX & " ;"
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                            sData = sLockPoint & " Z" & Str(GetCoordinates(sData, InStr(1, sData, "Z"), False) - dExtraX) & " ;"
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                            sData = "Z" & dsXpos & " ;"
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                            sData = "G00 X" & dsYpos & " ;"
                            PaintResult sData, mForm, mListbox, mResultPicture, False
                            sData = mListbox.List(mListbox.ListIndex)
                        End If
                    Loop Until InStr(1, sData, "G") <> 0
                    iLastReaded(iOpenFiles) = mListbox.ListIndex
                Case Else
            End Select
EndSelectCase:
        End If
    End If
    PaintResult sData, mForm, mListbox, mResultPicture, False
End Sub

Private Sub PaintResult(sData As String, mForm As Form, mListbox As ListBox, picResult As PictureBox, Optional PaintFast As Boolean = False)
    Dim blnRadius As Boolean
    Dim dFeed As Double
    If mForm!ChkDirect.Value = 1 Then
        PaintFast = True
    Else
        PaintFast = False
    End If
    If InStr(1, sData, ":") <> 0 Then
        [mForm].Caption = "NC Simulator O" & Mid(sData, 2, 4) & ".nc"
        'add for visible
        If Mid(sData, 2, 1) = "7" Then
            If mListbox.ListIndex = 0 And iOpenFiles = 1 Then
                picResult.ForeColor = QBColor(4)
                lLineColor = QBColor(4)
                xStartPosition = xWidth
                yStartPosition = yHeight
            Else
                GoTo SkipCode
            End If
        Else
        ' end add
        GoTo SkipCode
        'add for visible
        End If
        'end add
    End If
    If InStr(1, sData, "T") <> 0 Then
        If IsNumeric(GetCoordinates(sData, InStr(1, sData, "T"), False)) Then
        If InStr(1, sData, "G0") = 0 Then
            Dim lColorLong As Long
            Select Case intColor
                Case 0: intColor = 2: lColorLong = &HFF00FF
                Case 2: intColor = 3: lColorLong = &H80&
                Case 3: intColor = 4: lColorLong = &H80FF&
                Case 4: intColor = 5: lColorLong = &HC0&
                Case 5: intColor = 8: lColorLong = &HC000&
                Case 8: intColor = 11: lColorLong = &H8000&
                Case 11: intColor = 12: lColorLong = &HFF8080
                Case 12: intColor = 13: lColorLong = &HC00000
                Case 13: intColor = 2: lColorLong = &HFF00FF
            End Select
            'Debug.Print intColor
            picResult.ForeColor = lColorLong 'QBColor(intColor)
            lLineColor = lColorLong 'QBColor(intColor)
        End If
        End If
    End If
    If InStr(1, sData, "G0") <> 0 Then
        Dim iPos As Integer
        iPos = InStr(1, sData, "G0") + 2
        Select Case Mid(sData, iPos, 1)
            Case " ", "0"
                picResult.DrawStyle = 2 'ijlgang
                blnRadius = False
            Case "2", "3"
                picResult.DrawStyle = 0
                blnRadius = True
            Case Else
                picResult.DrawStyle = 0
                blnRadius = False
        End Select
    End If
    If InStr(1, sData, "G32") <> 0 Then
        picResult.DrawStyle = 0
        blnRadius = False
    End If
    On Error Resume Next
    If InStr(1, sData, "G20") <> 0 Then
        mForm!lblMeasureUnitV.Caption = "Inch"
    ElseIf InStr(1, sData, "G21") <> 0 Then
        mForm!lblMeasureUnitV.Caption = "MM"
    End If
    If InStr(1, sData, "G40") <> 0 Then
        mForm!lblRadCompV.Caption = "None"
    ElseIf InStr(1, sData, "G41") <> 0 Then
        mForm!lblRadCompV.Caption = "Left"
    ElseIf InStr(1, sData, "G42") <> 0 Then
        mForm!lblRadCompV.Caption = "Right"
    End If
    If InStr(1, sData, "G50 S") <> 0 Then
        mForm!lblMaxSpeedV.Caption = GetCoordinates(sData, InStr(1, sData, "S")) / dScaleFactor
    End If
    If InStr(1, sData, "G96") <> 0 Then
        mForm!lblConstant.Caption = "Constant CSS :"
        mForm!lblConstantV.Caption = GetCoordinates(sData, InStr(1, sData, "S")) / dScaleFactor
    ElseIf InStr(1, sData, "G97") <> 0 Then
        mForm!lblConstant.Caption = "Constant RPM :"
        mForm!lblConstantV.Caption = GetCoordinates(sData, InStr(1, sData, "S")) / dScaleFactor
    End If
    Static sFeedV As String
    If InStr(1, sData, "G98") <> 0 Then
        sFeedV = " mm/min"
        If mForm!lblMeasureUnitV.Caption = "Inch" Then
            sFeedV = " inch/min"
        End If
    ElseIf InStr(1, sData, "G99") <> 0 Then
        sFeedV = " mm/rev"
        If mForm!lblMeasureUnitV.Caption = "Inch" Then
            sFeedV = " inch/rev"
        End If
    End If
    If InStr(1, sData, "F") <> 0 Then
        dFeed = GetCoordinates(sData, InStr(1, sData, "F"), False)
        mForm!lblFeedV.Caption = Trim(dFeed) & sFeedV
    End If
    If InStr(1, sData, "M00") <> 0 Or sData = "(CLS) ;" Then
        Wait 2000
        frmGraphic.SetLabelsNothing
        DrawXZ_Axis picResult
        intColor = 0
        picResult.ForeColor = QBColor(intColor)
        lLineColor = QBColor(intColor)
        DoEvents
    ElseIf InStr(1, sData, "G10") <> 0 Then
        mForm!lblNullPoint.Caption = "Zero Point : Z" & GetCoordinates(sData, InStr(1, sData, "Z"), False)
        mForm!lblNullPoint.Visible = True
        Wait 2000
        mForm!lblNullPoint.Visible = False
    ElseIf InStr(1, sData, "G04") <> 0 Then
        Wait GetCoordinates(sData, InStr(1, sData, "U"), False)
    Else
        DrawContour sData, blnRadius, picResult, PaintFast
    End If
SkipCode:
    DoEvents
End Sub

Private Sub PaintForCalc(lstAutoCommands As ListBox, picResult As PictureBox, Optional PaintFast As Boolean = False)
    Dim iCount As Integer
    Dim blnRadius As Boolean
    Dim OldColor As Long
    OldColor = lLineColor
    lLineColor = vbBlack
    For iCount = 0 To lstAutoCommands.ListCount - 1
        lstAutoCommands.ListIndex = iCount
        If InStr(1, lstAutoCommands.List(iCount), "G0") <> 0 Then
            Dim iPos As Integer
            iPos = InStr(1, lstAutoCommands.List(iCount), "G0") + 2
            Select Case Mid(lstAutoCommands.List(iCount), iPos, 1)
                Case " ", "0"
                    picResult.DrawStyle = 2 'ijlgang
                    blnRadius = False
                Case "2", "3"
                    picResult.DrawStyle = 0
                    blnRadius = True
                Case Else
                    picResult.DrawStyle = 0
                    blnRadius = False
            End Select
        End If
        DrawContour lstAutoCommands.List(iCount), blnRadius, picResult, PaintFast
        DoEvents
    Next
    lLineColor = OldColor
End Sub

Private Function AddForCut(sData As String, dHowManyCuts As Double, dMaxX As Double, dMaxZ As Double, dMinX As Double, dMinZ As Double, iCut As Integer) As String
    Dim dStepX As Double
    Dim dStepZ As Double
    dStepX = (dMaxX - dMinX) / (IIf(dHowManyCuts = 1, 1, dHowManyCuts - 1))
    dStepZ = (dMaxZ - dMinZ) / (IIf(dHowManyCuts = 1, 1, dHowManyCuts - 1))
    If iCut = 1 Then
        dStepX = dMaxX
        dStepZ = dMaxZ
    ElseIf iCut = dHowManyCuts Then
        dStepX = dMinX
        dStepZ = dMinZ
    Else
        dStepX = dMaxX - (dStepX * (iCut - 1))
        dStepZ = dMaxZ - (dStepZ * (iCut - 1))
    End If
    If InStr(1, sData, "X") <> 0 Then
        sData = Left(sData, InStr(1, sData, "X")) & Trim(Str(GetCoordinates(sData, InStr(1, sData, "X"), False) + dStepX)) & Mid(sData, InStr(1, sData, "X") + LenCoordinates(sData, InStr(1, sData, "X"), False) + 1)
        'Debug.Print sData
    End If
    If InStr(1, sData, "Z") <> 0 Then
        sData = Left(sData, InStr(1, sData, "Z")) & Trim(Str(GetCoordinates(sData, InStr(1, sData, "Z"), False) + dStepZ)) & Mid(sData, InStr(1, sData, "Z") + LenCoordinates(sData, InStr(1, sData, "Z"), False) + 1)
        'Debug.Print sData
    End If
    If InStr(1, sData, "U") <> 0 Then
        sData = Left(sData, InStr(1, sData, "U")) & Trim(Str(GetCoordinates(sData, InStr(1, sData, "U"), False) + dStepX)) & Mid(sData, InStr(1, sData, "U") + LenCoordinates(sData, InStr(1, sData, "U"), False) + 1)
        'Debug.Print sData
    End If
    If InStr(1, sData, "W") <> 0 Then
        sData = Left(sData, InStr(1, sData, "W")) & Trim(Str(GetCoordinates(sData, InStr(1, sData, "W"), False) + dStepZ)) & Mid(sData, InStr(1, sData, "W") + LenCoordinates(sData, InStr(1, sData, "W"), False) + 1)
        'Debug.Print sData
    End If
    AddForCut = sData
End Function
