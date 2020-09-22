VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNC_Editor 
   Caption         =   "NC Editor"
   ClientHeight    =   5805
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4605
   Icon            =   "frmtNC_Editor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9128
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmtNC_Editor.frx":030A
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
   Begin VB.ListBox lstTemp 
      Height          =   5130
      Left            =   4560
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   9000
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Bestand openen"
      Filter          =   "Alle bestanden (*.*)|*.txt;*.nc|NC bestanden (*.nc)|*.nc|Tekst bestanden (*.txt)|*.txt"
      FilterIndex     =   2
   End
   Begin MSComctlLib.ImageList imlKnoppen 
      Left            =   9000
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":049C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":05AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":06C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":07D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":08E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":09F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":0B08
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":0C1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":0D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":0E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":0F50
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":1062
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":1174
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":1286
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":1398
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtNC_Editor.frx":14AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbKnoppen 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imlKnoppen"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Nieuw"
            Object.ToolTipText     =   "Nieuw"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Openen"
            Object.ToolTipText     =   "Openen"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Opslaan"
            Object.ToolTipText     =   "Opslaan"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Verwijderen"
            Object.ToolTipText     =   "Verwijderen"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Afdrukken"
            Object.ToolTipText     =   "Afdrukken"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Zoeken"
            Object.ToolTipText     =   "Zoeken"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Knippen"
            Object.ToolTipText     =   "Knippen"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Kopieren"
            Object.ToolTipText     =   "Kopïeren"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Plakken"
            Object.ToolTipText     =   "Plakken"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Ongedaan maken"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Object.ToolTipText     =   "Herhalen"
            ImageIndex      =   14
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuBestand 
      Caption         =   "&File"
      Begin VB.Menu mnuBestandNieuw 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuBestandOpenen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuBestandLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBestandOpslaan 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBestandOpslaanAls 
         Caption         =   "Sav&e As..."
      End
      Begin VB.Menu mnuBestandLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBestandVerwijderen 
         Caption         =   "Delete ..."
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuBestandLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBestandAfdrukken 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuBestandLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBestandAfsluiten 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuBewerken 
      Caption         =   "&Edit"
      Begin VB.Menu mnuBewerkenOngedaan 
         Caption         =   "Undo"
      End
      Begin VB.Menu mnuBewerkenHerhalen 
         Caption         =   "Redo"
      End
      Begin VB.Menu mnuBewerkenLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBewerkenKnippen 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuBewerkenKopieren 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuBewerkenPlakken 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuBewerkenLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBewerkenZoeken 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuBewerkenVervangen 
         Caption         =   "Replace"
      End
   End
   Begin VB.Menu mnuOpmaak 
      Caption         =   "Layout"
      Begin VB.Menu mnuOpmaakLettertype 
         Caption         =   "&Font..."
      End
      Begin VB.Menu mnuOpmaakLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpmaakUiteenTrekken 
         Caption         =   "Stretch Line"
      End
      Begin VB.Menu mnuOpmaakRegelNummersAan 
         Caption         =   "Add Line Numbers"
      End
      Begin VB.Menu mnuOpmaakRegelNummersUit 
         Caption         =   "Delete Line Numbers"
      End
   End
   Begin VB.Menu mnuTekst 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuTekstKnippen 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuTekstKopieren 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuTekstPlakken 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuTekstLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTekstRegelsUiteen 
         Caption         =   "Stretch Line"
      End
      Begin VB.Menu mnuTekstRegelNummerAan 
         Caption         =   "Add Line Numbers"
      End
      Begin VB.Menu mnuTekstRegelNummerUit 
         Caption         =   "Delete Line Numbers"
      End
   End
End
Attribute VB_Name = "frmNC_Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const AnInch As Long = 1440   '1440 twips per inch
Private Const QuarterInch As Long = 360

Dim sPathOpen As String
Dim blnChanged As Boolean

Dim blnIgnoreChanged As Boolean
Dim iChangedText As Integer
Dim iChangedMax As Integer
Dim sChangedText(1000) As String

Private Sub Form_Load()
    Dim PrintableWidth As Long
    Dim PrintableHeight As Long
    Dim x As Single

    'initialize the printer object
    x = Printer.TwipsPerPixelX
    Printer.Orientation = vbPRORPortrait  'vbPRORLandscape,vbPRORPortrait
    Printer.PrintQuality = vbPRPQLow 'vbPRPQDraft,vbPRPQLow,vbPRPQMedium,vbPRPQHigh
    Printer.ColorMode = vbPRCMMonochrome 'VbPRCMMonochrome,VbPRCMColor

    ' Tell the RTF to base it's display off of the printer
    'Call WYSIWYG_RTF(RichTextBox1, QuarterInch, QuarterInch, QuarterInch, QuarterInch, PrintableWidth, PrintableHeight) '1440 Twips=1 Inch

    ' Set the form width to match the line width
    'Me.Width = PrintableWidth + 200
    'Me.Height = PrintableHeight + 800
    
    If frmGraphic.dlgDialog.FileName = "" Then
        sPathOpen = App.Path & "\Nieuw"
        Me.Caption = "NC Editor : Nieuw"
        dlgDialog.InitDir = App.Path
    Else
        sPathOpen = frmGraphic.dlgDialog.FileName
        dlgDialog.FileName = sPathOpen
        dlgDialog.InitDir = GetPath(frmGraphic.dlgDialog.FileName)
        Me.Caption = "NC Editor : " & GetFileName(sPathOpen)
        iChangedText = 0
        RichTextBox1.LoadFile (sPathOpen), rtfText
        blnChanged = False
    End If
    FormOnTop Me.hWnd, True
End Sub

Private Sub Form_Resize()
   ' Position the RTF on form
   RichTextBox1.Move 30, tlbKnoppen.Height + 30, Me.ScaleWidth - 60, Me.ScaleHeight - 480
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmFind
    Unload frmReplace
    SaveChangedFile
    FormOnTop Me.hWnd, False
    'FormOnTop frmGraphic.hWnd, True
End Sub

Private Sub mnuBestandAfdrukken_Click()
    PrintRTF RichTextBox1, AnInch, AnInch, AnInch, AnInch
End Sub

Private Sub mnuBestandAfsluiten_Click()
    End
End Sub

Private Sub mnuBestandNieuw_Click()
    SaveChangedFile
    RichTextBox1.Text = ""
    iChangedText = 0
    sPathOpen = App.Path & "\Nieuw"
    Me.Caption = "NC Editor : " & GetFileName(sPathOpen)
    blnChanged = False
End Sub

Private Sub mnuBestandOpenen_Click()
    SaveChangedFile
    With dlgDialog
        .DialogTitle = "Bestand openen"
        .ShowOpen
        sPathOpen = .FileName
    End With
    If sPathOpen <> "" Then
        RichTextBox1.Text = ""
        iChangedText = 0
        RichTextBox1.LoadFile (sPathOpen), rtfText
        blnChanged = False
        Me.Caption = "NC Editor : " & GetFileName(sPathOpen)
    End If
End Sub

Private Sub mnuBestandOpslaan_Click()
    If GetFileName(sPathOpen) = "Nieuw" Or sPathOpen = "" Then
        mnuBestandOpslaanAls_Click
    ElseIf sPathOpen <> "" Then
        RichTextBox1.SaveFile sPathOpen, rtfText
        blnChanged = False
    End If
End Sub

Private Sub mnuBestandOpslaanAls_Click()
    With dlgDialog
        .DialogTitle = "Bestand opslaan"
        .ShowSave
        sPathOpen = .FileName
    End With
    If sPathOpen <> "" Then
        RichTextBox1.SaveFile sPathOpen, rtfText
        blnChanged = False
        Me.Caption = "NC Editor : " & GetFileName(sPathOpen)
    End If
End Sub

Private Sub mnuBestandVerwijderen_Click()
    If MsgBox(sPathOpen, vbOKCancel + vbQuestion, "Delete File ...") = vbOK Then
        Kill sPathOpen
    End If
End Sub

Private Sub mnuBewerkenHerhalen_Click()
    TextRedo
End Sub

Private Sub mnuBewerkenKnippen_Click()
    Knippen
End Sub

Private Sub mnuBewerkenKopieren_Click()
    Kopieren
End Sub

Private Sub mnuBewerkenOngedaan_Click()
    TextUndo
End Sub

Private Sub mnuBewerkenPlakken_Click()
    Plakken
End Sub

Private Sub mnuBewerkenVervangen_Click()
    FormOnTop Me.hWnd, False
    frmReplace.Combo1.AddItem RichTextBox1.SelText
    frmReplace.Combo1.ListIndex = frmReplace.Combo1.ListCount - 1
    frmReplace.Show
End Sub

Private Sub mnuBewerkenZoeken_Click()
    FormOnTop Me.hWnd, False
    frmFind.Combo1.AddItem RichTextBox1.SelText
    frmFind.Combo1.ListIndex = frmFind.Combo1.ListCount - 1
    frmFind.Show
End Sub

Private Sub mnuOpmaakLettertype_Click()
    LetterType
End Sub

Private Sub mnuOpmaakRegelNummersAan_Click()
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
    blnIgnoreChanged = True
    RichTextBox1.Text = ""
    blnIgnoreChanged = False
    RichTextBox1.LoadFile (sTempFile), rtfText
    blnChanged = True
    Kill sTempFile
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuOpmaakRegelNummersUit_Click()
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
    blnIgnoreChanged = True
    RichTextBox1.Text = ""
    blnIgnoreChanged = False
    RichTextBox1.LoadFile (sTempFile), rtfText
    blnChanged = True
    Kill sTempFile
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuOpmaakUiteenTrekken_Click()
    Dim FileNum As Integer
    Dim sTempFile As String
    Dim sData, sData1 As String
    Dim iCount As Integer
    Dim iLineCount As Integer
    Dim sChar As String
    sTempFile = App.Path & "\Temp.$$$"
    RichTextBox1.SaveFile sTempFile, rtfText
    lstTemp.Clear
    FileNum = FreeFile()
    Me.MousePointer = vbHourglass
    Open (sTempFile) For Input As #FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, sData
            iLineCount = iLineCount + 1
            sData1 = ""
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
    blnIgnoreChanged = True
    RichTextBox1.Text = ""
    blnIgnoreChanged = False
    RichTextBox1.LoadFile (sTempFile), rtfText
    blnChanged = True
    Kill sTempFile
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuTekstKnippen_Click()
    Knippen
End Sub

Private Sub mnuTekstKopieren_Click()
    Kopieren
End Sub

Private Sub mnuTekstPlakken_Click()
    Plakken
End Sub

Private Sub mnuTekstRegelNummerAan_Click()
    mnuOpmaakRegelNummersAan_Click
End Sub

Private Sub mnuTekstRegelNummerUit_Click()
    mnuOpmaakRegelNummersUit_Click
End Sub

Private Sub mnuTekstRegelsUiteen_Click()
    mnuOpmaakUiteenTrekken_Click
End Sub

Private Sub RichTextBox1_Change()
    blnChanged = True
    If Not blnIgnoreChanged Then
        iChangedText = iChangedText + 1
        iChangedMax = iChangedText
        sChangedText(iChangedText) = RichTextBox1.Text
    End If
End Sub

'Return the Filename part of the filespec
Private Function GetFileName(ByVal FileName As String) As String
    Dim L As Integer, j As Integer
    L = Len(FileName)
    For j = L To 1 Step -1
        If Mid(FileName, j, 1) = "\" Then Exit For
    Next j
    GetFileName = Mid(FileName, j + 1)
End Function

'Private Sub LoadFile(sOpenFile As String)
'    Dim FileNum As Integer
'    Dim sData As String
'    FileNum = FreeFile()
'    Open (sOpenFile) For Input As #FileNum
'        Do Until EOF(FileNum)
'            Line Input #FileNum, sData
'            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & sData
'        Loop
'    Close #FileNum
'End Sub

Private Sub LetterType()
    On Error Resume Next
    dlgDialog.Flags = cdlCFBoth + cdlCFEffects
    With RichTextBox1
        dlgDialog.FontName = .SelFontName
        dlgDialog.FontSize = .SelFontSize
        dlgDialog.FontBold = .SelBold
        dlgDialog.FontItalic = .SelItalic
        dlgDialog.FontUnderline = .SelUnderline
        dlgDialog.FontStrikethru = .SelStrikeThru
        dlgDialog.Color = .SelColor
    End With
    dlgDialog.ShowFont
    If Err.Number = 0 Then
        With RichTextBox1
            .SelFontName = dlgDialog.FontName
            .SelFontSize = dlgDialog.FontSize
            .SelBold = dlgDialog.FontBold
            .SelItalic = dlgDialog.FontItalic
            .SelUnderline = dlgDialog.FontUnderline
            .SelStrikeThru = dlgDialog.FontStrikethru
            .SelColor = dlgDialog.Color
        End With
    End If
End Sub

Private Sub Knippen()
    Clipboard.Clear
    Clipboard.SetText RichTextBox1.SelText, vbCFText
    RichTextBox1.SelText = ""
End Sub

Private Sub Kopieren()
    Clipboard.Clear
    Clipboard.SetText RichTextBox1.SelText, vbCFText
End Sub

Private Sub Plakken()
    RichTextBox1.SelText = Clipboard.GetText(vbCFText)
End Sub

Private Sub SaveChangedFile()
    If blnChanged = True Then
        Dim Question As String
        Question = """" & GetFileName(sPathOpen) & """" & " has changed" & vbCrLf & " save modifications."
        If MsgBox(Question, vbYesNo + vbQuestion, "NC Editor") = vbYes Then
            If GetFileName(sPathOpen) <> "Nieuw" Then
                mnuBestandOpslaan_Click
            Else
                mnuBestandOpslaanAls_Click
            End If
        End If
    End If
End Sub

Private Sub TextRedo()
    'This is the basic redo stuff.
    blnIgnoreChanged = True
    If iChangedMax > iChangedText Then
        iChangedText = iChangedText + 1
        On Error Resume Next
        RichTextBox1.Text = sChangedText(iChangedText)
    End If
    blnIgnoreChanged = False
End Sub

Private Sub TextUndo()
    'This says that if the Index is = to 0, then It shouldn't undo anymore
    If iChangedText = 0 Then Exit Sub
    
    'This is the basic undo stuff.
    blnIgnoreChanged = True
    iChangedText = iChangedText - 1
    On Error Resume Next
    RichTextBox1.Text = sChangedText(iChangedText)
    blnIgnoreChanged = False
End Sub

Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF10 Or KeyCode = vbKeyF12 Then
        MsgBox "Design by Patrick", vbOKOnly + vbInformation
    End If
End Sub

Private Sub RichTextBox1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuTekst
    End If
End Sub

Private Sub tlbKnoppen_ButtonClick(ByVal Button As MSComctlLib.Button)
    Button.MixedState = False
    Select Case Button.Key
        Case "Nieuw"
            mnuBestandNieuw_Click
        Case "Openen"
            mnuBestandOpenen_Click
        Case "Opslaan"
            mnuBestandOpslaan_Click
        Case "Afdrukken"
            mnuBestandAfdrukken_Click
        Case "Zoeken"
            mnuBewerkenZoeken_Click
        Case "Knippen"
            Knippen
        Case "Kopieren"
            Kopieren
        Case "Plakken"
            Plakken
        Case "Verwijderen"
            mnuBestandVerwijderen_Click
        Case "Undo"
            TextUndo
        Case "Redo"
            TextRedo
    End Select
End Sub

