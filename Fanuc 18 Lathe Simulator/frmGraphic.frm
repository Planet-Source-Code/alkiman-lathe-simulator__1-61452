VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmGraphic 
   Caption         =   "CNC Simulator"
   ClientHeight    =   8085
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10875
   FillStyle       =   0  'Solid
   Icon            =   "frmGraphic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   539
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   725
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstTemp 
      Height          =   255
      Left            =   11160
      TabIndex        =   38
      Top             =   6120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5520
      Top             =   4440
   End
   Begin VB.PictureBox picViewCommand 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   900
      Index           =   2
      Left            =   3480
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   35
      ToolTipText     =   "Start simulation"
      Top             =   6840
      Width           =   900
   End
   Begin VB.PictureBox picViewCommand 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   720
      Index           =   1
      Left            =   4080
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   46
      TabIndex        =   33
      ToolTipText     =   "Coolant On/Off"
      Top             =   6000
      Width           =   720
   End
   Begin VB.PictureBox picViewCommand 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   720
      Index           =   0
      Left            =   3120
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   46
      TabIndex        =   32
      ToolTipText     =   "Turn Direction"
      Top             =   6000
      Width           =   720
   End
   Begin VB.CheckBox ChkDirect 
      Caption         =   "Direct"
      Height          =   375
      Left            =   8640
      TabIndex        =   29
      Top             =   7200
      Width           =   1215
   End
   Begin VB.ListBox lstAutomatic 
      Height          =   1620
      Left            =   5040
      TabIndex        =   28
      Top             =   6000
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   16
      Top             =   5880
      Width           =   2775
      Begin VB.Label lblNullPoint 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1215
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblMaxSpeed 
         Caption         =   "Max Spindle Speed :"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblMaxSpeedV 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1800
         TabIndex        =   25
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblRadComp 
         Caption         =   "Radius Compensation :"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblRadCompV 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblMeasureUnitV 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblMeasureUnit 
         Caption         =   "Inch/Metric :"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblConstantV 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1890
         TabIndex        =   20
         Top             =   960
         Width           =   675
      End
      Begin VB.Label lblConstant 
         Caption         =   "Constant :"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblFeedV 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblFeed 
         Caption         =   "Feedrate :"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.ListBox lstCommands 
      Height          =   1620
      Left            =   5040
      TabIndex        =   15
      Top             =   6000
      Width           =   3375
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   10
      Left            =   8640
      Max             =   0
      Min             =   20
      TabIndex        =   12
      Top             =   6840
      Value           =   10
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   8640
      Max             =   200
      Min             =   10
      TabIndex        =   9
      Top             =   6240
      Value           =   10
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   6600
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Alle bestanden (*.*)|*.*|NC bestanden (*.nc)|*.nc|Tekst bestanden (*.txt)|*.txt"
      FilterIndex     =   2
   End
   Begin VB.CommandButton cmdLineUpRight 
      Height          =   495
      Left            =   13200
      Picture         =   "frmGraphic.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton cmdLineDownLeft 
      Height          =   495
      Left            =   12000
      Picture         =   "frmGraphic.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton cmdLineDownRight 
      Height          =   495
      Left            =   13200
      Picture         =   "frmGraphic.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton cmdLineRight 
      Height          =   495
      Left            =   13200
      Picture         =   "frmGraphic.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6720
      Width           =   495
   End
   Begin VB.CommandButton cmdLineDown 
      Height          =   495
      Left            =   12600
      Picture         =   "frmGraphic.frx":1412
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton cmdLineUp 
      Height          =   495
      Left            =   12600
      Picture         =   "frmGraphic.frx":1854
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton cmdLineUpLeft 
      Height          =   495
      Left            =   12000
      Picture         =   "frmGraphic.frx":1C96
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton cmdLineLeft 
      Height          =   495
      Left            =   12000
      Picture         =   "frmGraphic.frx":20D8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   495
   End
   Begin VB.PictureBox picSave 
      AutoRedraw      =   -1  'True
      ForeColor       =   &H00008000&
      Height          =   5595
      Left            =   135
      ScaleHeight     =   369
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   702
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   10590
   End
   Begin VB.PictureBox picCalculate 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00008000&
      Height          =   5595
      Left            =   135
      ScaleHeight     =   369
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   702
      TabIndex        =   27
      Top             =   150
      Visible         =   0   'False
      Width           =   10590
   End
   Begin VB.PictureBox picZoom 
      AutoRedraw      =   -1  'True
      Height          =   5595
      Left            =   135
      ScaleHeight     =   369
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   702
      TabIndex        =   36
      Top             =   150
      Visible         =   0   'False
      Width           =   10590
   End
   Begin VB.PictureBox picResult 
      AutoRedraw      =   -1  'True
      ForeColor       =   &H00008000&
      Height          =   5595
      Left            =   135
      ScaleHeight     =   369
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   702
      TabIndex        =   0
      Top             =   150
      Width           =   10590
      Begin VB.Shape SelectShape 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         FillColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1080
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image imgSelect 
         Height          =   615
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.PictureBox picCommands 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3630
      Left            =   480
      Picture         =   "frmGraphic.frx":251A
      ScaleHeight     =   3630
      ScaleWidth      =   7155
      TabIndex        =   34
      Top             =   225
      Width           =   7155
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5595
      Left            =   10920
      TabIndex        =   37
      Top             =   150
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   9869
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmGraphic.frx":56F0C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblDrawSpeed 
      Caption         =   "Draw speed"
      Height          =   255
      Left            =   8640
      TabIndex        =   14
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      Caption         =   "10"
      Height          =   255
      Left            =   10200
      TabIndex        =   13
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label lblZoomFactor 
      Caption         =   "Zoom factor"
      Height          =   255
      Left            =   8640
      TabIndex        =   11
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label lblZoom 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Left            =   10200
      TabIndex        =   10
      Top             =   6240
      Width           =   375
   End
   Begin VB.Menu mnuBestand 
      Caption         =   "&File"
      Begin VB.Menu mnuBestandOpenen 
         Caption         =   "Open ..."
      End
      Begin VB.Menu mnuBestandLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBestandEditor 
         Caption         =   "Editor ..."
      End
      Begin VB.Menu mnuBestandLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBestandSavePicture 
         Caption         =   "Save Picture"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBestandLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBestandAfsluiten 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuBewerking 
      Caption         =   "&Simulation"
      Begin VB.Menu mnuBewerkingStarten 
         Caption         =   "Start"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBewerkingLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBewerkingVernieuw 
         Caption         =   "Restart"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBewerkingLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBewerkingOptioneel 
         Caption         =   "Optional stop"
      End
      Begin VB.Menu mnuBewerkingStap 
         Caption         =   "Step by step"
      End
   End
   Begin VB.Menu mnuOntwerp 
      Caption         =   "Design"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuOntwerpSimulatie 
         Caption         =   "Simulation"
      End
      Begin VB.Menu mnuOntwerpLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOntwerpRegelOn 
         Caption         =   "Add Line Numbers"
      End
      Begin VB.Menu mnuOntwerpRegelOff 
         Caption         =   "Delete Line Numbers"
      End
      Begin VB.Menu mnuOntwerpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOntwerpNewTool 
         Caption         =   "New Tool ..."
      End
      Begin VB.Menu mnuOntwerpLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOntwerpG9094 
         Caption         =   "G90 - G92 - G94"
      End
      Begin VB.Menu mnuOntwerpLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOntwerpEinde 
         Caption         =   "M30 End program"
      End
      Begin VB.Menu mnuOntwerpLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOntwerpClear 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "frmGraphic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOACTIVATE = &H10
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWid6th As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Dim blnRedraw As Boolean
Dim blnDrawing As Boolean
Dim blnRunning As Boolean
Dim iTurn As String

Private Type POINT
    x As Long
    y As Long
End Type

Dim sPoint As POINT
Dim ePoint As POINT

Public blnSetRegion As Boolean

Dim SelectWidth, SelectHeight, SelectTop, SelectLeft As Double

Private Sub cmdLineDown_Click()
    MsgBox "Only available by licensed version", vbOKOnly + vbInformation
End Sub

Private Sub cmdLineDownLeft_Click()
    MsgBox "Only available by licensed version", vbOKOnly + vbInformation
End Sub

Private Sub cmdLineDownRight_Click()
    MsgBox "Only available by licensed version", vbOKOnly + vbInformation
End Sub

Private Sub cmdLineLeft_Click()
    MsgBox "Only available by licensed version", vbOKOnly + vbInformation
End Sub

Private Sub cmdLineRight_Click()
    MsgBox "Only available by licensed version", vbOKOnly + vbInformation
End Sub

Private Sub cmdLineUp_Click()
    MsgBox "Only available by licensed version", vbOKOnly + vbInformation
End Sub

Private Sub cmdLineUpLeft_Click()
    MsgBox "Only available by licensed version", vbOKOnly + vbInformation
End Sub

Private Sub cmdLineUpRight_Click()
    MsgBox "Only available by licensed version", vbOKOnly + vbInformation
End Sub

Private Sub Form_Load()
    Dim ScreenResolution As String
    ScreenResolution = Screen.Width / Screen.TwipsPerPixelX & "x" & Screen.Height / Screen.TwipsPerPixelY
    If ScreenResolution <> "1024x768" Then
        If MsgBox("Set Screen Resolution At 1024x768", vbQuestion + vbYesNo, "CNC Simulator : Screen Resolution " & Screen.Width / Screen.TwipsPerPixelX & "x" & Screen.Height / Screen.TwipsPerPixelY) = vbYes Then
            End
        End If
    End If
    'SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    'FormOnTop Me.hWnd, True
    HScroll2.Value = 0
    HScroll1.Value = 34.5
    xWidth = 660
    yHeight = 340
    DrawXZ_Axis picResult
    Me.Refresh
    imgSelect.Top = picResult.ScaleTop
    imgSelect.Left = picResult.ScaleLeft
    imgSelect.Width = picResult.Width
    imgSelect.Height = picResult.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    Unload frmTool
    Unload frmG9094
    Unload frmNC_Editor
    'FormOnTop Me.hWnd, False
    End
End Sub

Private Sub HScroll1_Change()
    lblZoom.Caption = HScroll1.Value / 10
    dScaleFactor = HScroll1.Value / 10
End Sub

Private Sub HScroll1_Scroll()
    lblZoom.Caption = HScroll1.Value / 10
    dScaleFactor = HScroll1.Value / 10
End Sub

Private Sub HScroll2_Change()
    lblSpeed.Caption = (20 - HScroll2.Value) * 5
    lTimeMilliSeconds = HScroll2.Value
End Sub

Private Sub HScroll2_Scroll()
    lblSpeed.Caption = (20 - HScroll2.Value) * 5
    lTimeMilliSeconds = HScroll2.Value
End Sub

Private Sub imgSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If blnSetRegion = False Then
        blnSetRegion = True
        sPoint.x = x
        sPoint.y = y
    End If
End Sub

Private Sub imgSelect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If blnSetRegion = True And Button <> 0 Then
        ePoint.x = x
        ePoint.y = y
        SelectTop = IIf(sPoint.y <= ePoint.y, sPoint.y, ePoint.y) / 15
        SelectLeft = IIf(sPoint.x <= ePoint.x, sPoint.x, ePoint.x) / 15
        SelectHeight = IIf(sPoint.y >= ePoint.y, sPoint.y - ePoint.y, ePoint.y - sPoint.y) / 15
        SelectWidth = IIf(sPoint.x >= ePoint.x, sPoint.x - ePoint.x, ePoint.x - sPoint.x) / 15
        SelectShape.Left = SelectLeft
        SelectShape.Top = SelectTop
        SelectShape.Width = SelectWidth
        SelectShape.Height = SelectHeight
        SelectShape.Visible = True
    Else
        ePoint.x = x
        ePoint.y = y
        blnSetRegion = False
    End If
End Sub

Private Sub imgSelect_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo errHandle
    SelectShape.Visible = False
    If blnSetRegion = True Then
        If sPoint.x <> ePoint.x And sPoint.y <> ePoint.y Then
            Dim intZoomPart As Integer
            If (picResult.ScaleWidth / SelectWidth) >= (picResult.ScaleHeight / SelectHeight) Then
                intZoomPart = Int(picResult.ScaleHeight / SelectHeight)
            Else
                intZoomPart = Int(picResult.ScaleWidth / SelectWidth)
            End If
            Set picZoom.Picture = Nothing
            StretchBlt picZoom.hdc, 0, 0, SelectWidth * intZoomPart, SelectHeight * intZoomPart, picResult.hdc, SelectLeft, SelectTop, SelectWidth, SelectHeight, SRCCOPY
            picZoom.Visible = True
        End If
    End If
errHandle:
    blnSetRegion = False
End Sub

Private Sub mnuBestandAfsluiten_Click()
    End
End Sub

Private Sub mnuBestandEditor_Click()
    'FormOnTop Me.hWnd, False
    frmNC_Editor.dlgDialog.FileName = dlgDialog.FileName
    frmNC_Editor.Show
End Sub

Private Sub mnuBestandOpenen_Click()
    Dim sPathOpen() As String
    Dim FileNum As Integer
    Dim iRepeat() As Integer
    Dim iRunning As Integer
    Dim iLineCounter As Integer
    Dim sData As String
    If blnDrawing = True Then Exit Sub
    blnDrawing = True
    FileNum = FreeFile()
    iOpenFiles = iOpenFiles + 1
    ReDim Preserve sPathOpen(iOpenFiles)
    ReDim Preserve iRepeat(iOpenFiles)
    ReDim Preserve iLastReaded(iOpenFiles)
    iLastReaded(iOpenFiles) = 0
    iRepeat(iOpenFiles) = 1
    If blnRedraw = False Then
        With dlgDialog
            .DialogTitle = "Bestand openen"
            .ShowOpen
            sPathOpen(iOpenFiles) = .FileName
        End With
    Else
        sPathOpen(iOpenFiles) = dlgDialog.FileName
    End If
    SetLabelsNothing
    If sPathOpen(iOpenFiles) = "" Then
        iOpenFiles = 0
        blnDrawing = False
        Exit Sub
    End If
    mnuBewerkingVernieuw.Enabled = False
    mnuBestandSavePicture.Enabled = False
    mnuBewerkingStarten.Enabled = True
    picViewCommand(2).Enabled = True
    DrawXZ_Axis picResult
    GoTo OpenMasterFile

OpenSubFile:
    iOpenFiles = iOpenFiles + 1
    ReDim Preserve sPathOpen(iOpenFiles)
    ReDim Preserve iRepeat(iOpenFiles)
    ReDim Preserve iLastReaded(iOpenFiles)
    iRepeat(iOpenFiles) = Mid(lstCommands.List(lstCommands.ListIndex), InStr(1, lstCommands.List(lstCommands.ListIndex), "P") + 1, LenCoordinates(lstCommands.List(lstCommands.ListIndex), InStr(1, lstCommands.List(lstCommands.ListIndex), "P"), False) - 4)
    sPathOpen(iOpenFiles) = GetPath(dlgDialog.FileName) & "O" & Right(Mid(lstCommands.List(lstCommands.ListIndex), InStr(1, lstCommands.List(lstCommands.ListIndex), "P") + 1, LenCoordinates(lstCommands.List(lstCommands.ListIndex), InStr(1, lstCommands.List(lstCommands.ListIndex), "P"), False)), 4) & ".nc"
OpenMasterFile:
    For iRunning = 1 To iRepeat(iOpenFiles)
        If sPathOpen(iOpenFiles) <> "" Then

ContinuePrevious:
            If iRunning = 1 Then
                ReDim Preserve iRepeated(iOpenFiles)
                iRepeated(iOpenFiles) = iRunning
                Open (sPathOpen(iOpenFiles)) For Input As #FileNum
                    lstCommands.Clear
                    Do Until EOF(FileNum)
                        Line Input #FileNum, sData
                        If sData = "&HE:%" Then
                            Line Input #FileNum, sData
                        End If
                        If sData = "%" Or sData = "" Then
                            Exit Do
                        Else
                            lstCommands.AddItem CheckCommand(sData)
                            If InStr(1, sData, ":") <> 0 Then
                                Me.Caption = "NC Simulator O" & Mid(sData, 2, 4) & ".nc"
                            End If
                        End If
                    Loop
                Close #FileNum
                If lstCommands.ListCount = iLastReaded(iOpenFiles) Then
                    Exit For
                End If
                lstCommands.ListIndex = iLastReaded(iOpenFiles)
                lstCommands.TopIndex = lstCommands.ListIndex
            Else
                lstCommands.ListIndex = 0
                iLastReaded(iOpenFiles) = 0
            End If
            Do
                lstCommands.ListIndex = iLastReaded(iOpenFiles)
                lstCommands.TopIndex = lstCommands.ListIndex
                sData = lstCommands.List(lstCommands.ListIndex)
                StopByLine sData
                ViewCommand sData
                If InStr(1, sData, "M98") = 0 Then
                    DoCommandFanuc18TC Me, lstCommands, lstCommands.ListIndex, picResult, picCalculate
                End If
                If InStr(1, lstCommands.List(lstCommands.ListIndex), "M98") <> 0 Then
                    If FileExist(GetPath(dlgDialog.FileName) & "O" & Right(Mid(lstCommands.List(lstCommands.ListIndex), InStr(1, lstCommands.List(lstCommands.ListIndex), "P") + 1, LenCoordinates(lstCommands.List(lstCommands.ListIndex), InStr(1, lstCommands.List(lstCommands.ListIndex), "P"), False)), 4) & ".nc") Then
                        iLastReaded(iOpenFiles) = iLastReaded(iOpenFiles) + 1
                        GoTo OpenSubFile
                    Else
                        MsgBox GetPath(dlgDialog.FileName) & "O" & Right(GetCoordinates(lstCommands.List(lstCommands.ListIndex), InStr(1, lstCommands.List(lstCommands.ListIndex), "P"), False), 4) & ".nc", vbOKOnly + vbInformation, "File Not Find"
                    End If
                End If
                iLastReaded(iOpenFiles) = iLastReaded(iOpenFiles) + 1
            Loop Until lstCommands.ListIndex = lstCommands.ListCount - 1
        End If
        If iRunning = iRepeat(iOpenFiles) Then Exit For
    Next
    iOpenFiles = iOpenFiles - 1
    If iOpenFiles > 0 Then
        ReDim Preserve sPathOpen(iOpenFiles)
        ReDim Preserve iRepeat(iOpenFiles)
        ReDim Preserve iRepeated(iOpenFiles)
        ReDim Preserve iLastReaded(iOpenFiles)
        iLineCounter = 0
        iRunning = iRepeated(iOpenFiles)
        GoTo ContinuePrevious
    End If
    iLastReaded(iOpenFiles) = 0
    blnDrawing = False
    mnuBewerkingVernieuw.Enabled = True
    mnuBestandSavePicture.Enabled = True
    mnuBewerkingStarten.Enabled = False
    lblNullPoint.Caption = "Program done"
    lblNullPoint.Visible = True
    Wait 2000
    lblNullPoint.Visible = False
    If InStr(1, sData, "M30") <> 0 Then lstCommands.ListIndex = 0
End Sub

Private Sub mnuBestandSavePicture_Click()
    Dim DrawSpeed As Double
    Dim blnStap As Boolean
    Dim blnOptional As Boolean
    DrawSpeed = lTimeMilliSeconds
    blnStap = mnuBewerkingStap.Checked
    blnOptional = mnuBewerkingOptioneel.Checked
    mnuBewerkingStap.Checked = False
    mnuBewerkingOptioneel.Checked = False
    Set picSave.Picture = picResult.Image
    picSave.Visible = True
    picResult.BackColor = vbWhite
    picResult.Refresh
    HScroll2.Value = 0
    picViewCommand_Click (2)
    SavePicture picResult.Image, App.Path & "\" & Mid(Me.Caption, 13, Len(Me.Caption) - 15) & ".bmp"
    HScroll2.Value = DrawSpeed
    picResult.BackColor = &H8000000F
    picResult.Refresh
    BitBlt picResult.hdc, 0, 0, picResult.ScaleWidth, picResult.ScaleHeight, picSave.hdc, 0, 0, SRCCOPY
    picSave.Visible = False
    Dim sMessage As String
    sMessage = App.Path & "\" & Mid(Me.Caption, 13, Len(Me.Caption) - 15) & ".bmp"
    MsgBox sMessage, vbOKOnly + vbInformation, "Save Picture As ..."
End Sub

Private Sub mnuBewerkingOptioneel_Click()
    If mnuBewerkingOptioneel.Checked = False Then
        mnuBewerkingOptioneel.Checked = True
    Else
        mnuBewerkingOptioneel.Checked = False
    End If
End Sub

Private Sub mnuBewerkingStap_Click()
    If mnuBewerkingStap.Checked = False Then
        mnuBewerkingStap.Checked = True
    Else
        mnuBewerkingStap.Checked = False
    End If
End Sub

Private Sub mnuBewerkingStarten_Click()
    picViewCommand_Click (2)
End Sub

Private Sub mnuBewerkingVernieuw_Click()
    blnRedraw = True
    mnuBestandOpenen_Click
    blnRedraw = False
End Sub

Private Sub mnuOntwerpClear_Click()
    RichTextBox1.Text = ""
End Sub

Private Sub mnuOntwerpEinde_Click()
    Dim strTool As String
    strTool = strTool & "G00 X250 Z250 T0 M09 ;" & vbCrLf
    strTool = strTool & "M30 ;" & vbCrLf
    strTool = strTool & "%"
    Clipboard.Clear
    Clipboard.SetText strTool, vbCFText
    RichTextBox1.SelText = Clipboard.GetText(vbCFText)
End Sub

Private Sub mnuOntwerpG9094_Click()
    frmG9094.Show vbModal
End Sub

Private Sub mnuOntwerpNewTool_Click()
    frmTool.Show vbModal
End Sub

Private Sub mnuOntwerpRegelOff_Click()
    Dim FileNum As Integer
    Dim sTempFile As String
    Dim sData, sData1 As String
    Dim iCount As Integer
    Dim iLineCount As Integer
    sTempFile = App.Path & "\Temp.$$$"
    RichTextBox1.SaveFile sTempFile, rtfText
    lstTemp.Clear
    FileNum = FreeFile()
    Me.MousePointer = vbHourglass
    Open (sTempFile) For Input As #FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, sData
            iLineCount = iLineCount + 1
            Select Case iLineCount
                Case 1
                    sData1 = "&HE:%"
                Case 2
                    sData1 = sData
                Case Else
                    If Left(sData, 1) = "N" Then
                        iCount = InStr(1, sData, " ")
                        sData1 = Mid(sData, iCount + 1)
                    Else
                        sData1 = sData
                    End If
            End Select
            lstTemp.AddItem sData1
        Loop
    lstTemp.List(0) = "&HE:%"
    lstTemp.List(iLineCount - 1) = "%"
    Close #FileNum
    FileNum = FreeFile()
    Open (sTempFile) For Output As #FileNum
        For iCount = 0 To lstTemp.ListCount - 1
            Print #FileNum, lstTemp.List(iCount)
        Next
    Close #FileNum
    RichTextBox1.Text = ""
    RichTextBox1.LoadFile (sTempFile), rtfText
    Kill sTempFile
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuOntwerpRegelOn_Click()
    Dim FileNum As Integer
    Dim sTempFile As String
    Dim sData, sData1 As String
    Dim iCount As Integer
    Dim iLineCount As Integer
    sTempFile = App.Path & "\Temp.$$$"
    RichTextBox1.SaveFile sTempFile, rtfText
    lstTemp.Clear
    FileNum = FreeFile()
    Me.MousePointer = vbHourglass
    Open (sTempFile) For Input As #FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, sData
            iLineCount = iLineCount + 1
            Select Case iLineCount
                Case 1
                    sData1 = "&HE:%"
                Case 2
                    sData1 = sData
                Case Else
                    If Left(sData, 1) <> "N" Then
                        iCount = (iLineCount - 2) * 5
                        sData1 = "N" & iCount & " " & sData
                    Else
                        sData1 = sData
                    End If
            End Select
            lstTemp.AddItem sData1
        Loop
    lstTemp.List(0) = "&HE:%"
    lstTemp.List(iLineCount - 1) = "%"
    Close #FileNum
    FileNum = FreeFile()
    Open (sTempFile) For Output As #FileNum
        For iCount = 0 To lstTemp.ListCount - 1
            Print #FileNum, lstTemp.List(iCount)
        Next
    Close #FileNum
    RichTextBox1.Text = ""
    RichTextBox1.LoadFile (sTempFile), rtfText
    Kill sTempFile
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuOntwerpSimulatie_Click()
    Dim sPathOpen As String
    sPathOpen = App.Path & "\T0000.Nc"
    RichTextBox1.SaveFile sPathOpen, rtfText
    dlgDialog.FileName = sPathOpen
    blnRunning = True
    mnuBewerkingVernieuw_Click
End Sub

Private Sub ViewCommand(sData As String)
    If InStr(1, sData, "M03") <> 0 Then
        iTurn = "M03"
        BitBlt picViewCommand(0).hdc, 0, 0, 47, 47, picCommands.hdc, 145, 4, SRCCOPY
    ElseIf InStr(1, sData, "M04") <> 0 Then
        iTurn = "M04"
        BitBlt picViewCommand(0).hdc, 0, 0, 47, 47, picCommands.hdc, 98, 4, SRCCOPY
    ElseIf InStr(1, sData, "M08") <> 0 Then
        BitBlt picViewCommand(1).hdc, 0, 0, 47, 47, picCommands.hdc, 51, 51, SRCCOPY
    ElseIf InStr(1, sData, "M09") <> 0 Then
        BitBlt picViewCommand(1).hdc, 0, 0, 47, 47, picCommands.hdc, 4, 51, SRCCOPY
    End If
    If InStr(1, sData, "M30") <> 0 Then
        If iTurn = "M03" Then
            BitBlt picViewCommand(0).hdc, 0, 0, 47, 47, picCommands.hdc, 51, 4, SRCCOPY
        Else
            BitBlt picViewCommand(0).hdc, 0, 0, 47, 47, picCommands.hdc, 4, 4, SRCCOPY
        End If
        If iOpenFiles <= 1 Then
            StretchBlt picViewCommand(2).hdc, 0, 0, 60, 60, picCommands.hdc, 98, 51, 46, 46, SRCCOPY
            blnRunning = False
        End If
    End If
    Dim i As Integer
    For i = 0 To 2
        picViewCommand(i).Refresh
    Next
End Sub

Private Sub picViewCommand_Click(Index As Integer)
    If Index <> 2 Then Exit Sub
    If mnuBewerkingVernieuw.Enabled = True And blnDrawing = False Then
        blnRunning = True
        StretchBlt picViewCommand(2).hdc, 0, 0, 60, 60, picCommands.hdc, 192, 51, 46, 46, SRCCOPY
        mnuBewerkingVernieuw_Click
        StretchBlt picViewCommand(2).hdc, 0, 0, 60, 60, picCommands.hdc, 98, 51, 46, 46, SRCCOPY
        blnRunning = False
    Else
        blnRunning = True
        StretchBlt picViewCommand(2).hdc, 0, 0, 60, 60, picCommands.hdc, 192, 51, 46, 46, SRCCOPY
    End If
    picViewCommand(2).Refresh
End Sub

Private Sub picZoom_Click()
    picZoom.Visible = False
End Sub

Private Sub RichTextBox1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuOntwerp
    End If
End Sub

Private Sub Timer1_Timer()
    ViewCommand "M03"
    ViewCommand "M30 M09"
    Timer1.Enabled = False
End Sub

Public Sub StopByLine(sData As String)
    Static blnDone As Boolean
    Do
        If blnDone = False Then
            If InStr(1, sData, "M01") <> 0 And mnuBewerkingOptioneel.Checked = True Then
                blnRunning = False
                StretchBlt picViewCommand(2).hdc, 0, 0, 60, 60, picCommands.hdc, 98, 51, 46, 46, SRCCOPY
                If iTurn = "M03" Then
                    BitBlt picViewCommand(0).hdc, 0, 0, 47, 47, picCommands.hdc, 51, 4, SRCCOPY
                Else
                    BitBlt picViewCommand(0).hdc, 0, 0, 47, 47, picCommands.hdc, 4, 4, SRCCOPY
                End If
            ElseIf mnuBewerkingStap.Checked = True Then
                blnRunning = False
                StretchBlt picViewCommand(2).hdc, 0, 0, 60, 60, picCommands.hdc, 145, 51, 46, 46, SRCCOPY
            End If
            blnDone = True
            picViewCommand(0).Refresh
            picViewCommand(2).Refresh
        End If
        DoEvents
    Loop Until blnRunning = True
    blnDone = False
    If InStr(1, sData, "G32") <> 0 Then
        If InStr(1, sData, ":") <> 0 Then
            blnDone = False
        Else
            blnDone = True
        End If
    End If
End Sub

Public Sub SetLabelsNothing()
    lblConstantV = ""
    lblFeedV = ""
    lblMaxSpeedV = ""
    lblMeasureUnitV = ""
    lblRadCompV = ""
End Sub

