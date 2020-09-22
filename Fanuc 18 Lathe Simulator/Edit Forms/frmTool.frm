VERSION 5.00
Begin VB.Form frmTool 
   Caption         =   "New Tool"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4200
   Icon            =   "frmTool.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   4200
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2400
      TabIndex        =   22
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   495
      Left            =   480
      TabIndex        =   21
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Frame Frame7 
      Caption         =   "Spindle Turns"
      Height          =   855
      Left            =   2430
      TabIndex        =   18
      Top             =   2520
      Width           =   1620
      Begin VB.OptionButton Option8 
         Caption         =   "M03 - Right"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option9 
         Caption         =   "M04 - Left"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   510
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "G96 - G97"
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   2340
      Begin VB.TextBox txtCutSpeed 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Text            =   "200"
         Top             =   330
         Width           =   615
      End
      Begin VB.OptionButton Option7 
         Caption         =   "RPM"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   510
         Width           =   1200
      End
      Begin VB.OptionButton Option6 
         Caption         =   "CSS"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1200
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "T-Tool : Place - Correction - Comment"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   3930
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   1920
         TabIndex        =   11
         ToolTipText     =   "Beschrijving van bewerking - gereedschap"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtToolCorrection 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Text            =   "01"
         ToolTipText     =   "Nummer van 1 tot 99"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtTool 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Text            =   "01"
         ToolTipText     =   "Nummer van 1 tot 99"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "G98 - G99"
      Height          =   855
      Left            =   2715
      TabIndex        =   5
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton Option5 
         Caption         =   "mm/rev"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   510
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "mm/min"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "G40 Radius"
      Height          =   855
      Left            =   1410
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton Option3 
         Caption         =   "None"
         Height          =   255
         Left            =   200
         TabIndex        =   4
         Top             =   510
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "compensation"
         Height          =   255
         Left            =   200
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "G50 Maximum Spindle Speed"
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   3930
      Begin VB.TextBox txtMaxTour 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Text            =   "3000"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "G20 - G21"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton Option1 
         Caption         =   "Inch"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Metric"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim strTool As String
    If frmGraphic.RichTextBox1.Text = "" Then
        strTool = "&HE:%" & vbCrLf
        strTool = strTool & ":0000 ;" & vbCrLf
        strTool = strTool & IIf(Option1 = True, "G20 ", "G21 ") & "G40 " & IIf(Option4 = True, "G98 ;", "G99 ;") & vbCrLf
        strTool = strTool & "G00 X250 Z250 T0 ;" & vbCrLf
    Else
        strTool = strTool & "G00 X250 Z250 T0 M09 ;" & vbCrLf
        strTool = strTool & "M01 ;" & vbCrLf
    End If
    If txtDescription <> "" Then
        strTool = strTool & "(" & txtDescription & ") ;" & vbCrLf
    End If
    strTool = strTool & "T" & txtTool & txtToolCorrection & " ;" & vbCrLf
    strTool = strTool & "G50 S" & Trim(Str(txtMaxTour)) & " ;" & vbCrLf
    strTool = strTool & IIf(Option6 = True, "G96 S", "G97 S") & Trim(Str(txtCutSpeed)) & IIf(Option8 = True, " M03 ;", " M04 ;") & vbCrLf
    Clipboard.Clear
    Clipboard.SetText strTool, vbCFText
    frmGraphic.RichTextBox1.SelText = Clipboard.GetText(vbCFText)
    'Debug.Print strTool
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = frmGraphic.Top + 945
    Me.Left = frmGraphic.Left + 6420
End Sub

Private Sub Option1_Click()
    Option4.Caption = "inch/min"
    Option5.Caption = "inch/rev"
End Sub

Private Sub Option2_Click()
    Option4.Caption = "mm/min"
    Option5.Caption = "mm/rev"
End Sub

Private Sub Option6_LostFocus()
    txtCutSpeed = 200
End Sub

Private Sub Option7_LostFocus()
    txtCutSpeed = 700
End Sub

Private Sub txtMaxTour_Change()
    If Not IsNumeric(txtMaxTour.Text) Then
        txtMaxTour.Text = ""
    End If
End Sub

Private Sub txtMaxTour_GotFocus()
    txtMaxTour.SelStart = 0
    txtMaxTour.SelLength = Len(txtMaxTour.Text)
End Sub

Private Sub txtMaxTour_Validate(Cancel As Boolean)
    If txtMaxTour > 5000 Or txtMaxTour < 100 Then
        Cancel = True
        txtMaxTour.SelStart = 0
        txtMaxTour.SelLength = Len(txtMaxTour.Text)
    End If
End Sub

Private Sub txtTool_Change()
    If Not IsNumeric(txtTool.Text) Then
        txtTool.Text = ""
    End If
End Sub

Private Sub txtTool_GotFocus()
    txtTool.SelStart = 0
    txtTool.SelLength = Len(txtTool.Text)
End Sub

Private Sub txtTool_Validate(Cancel As Boolean)
    If txtTool > 99 Or txtTool < 1 Then
        Cancel = True
        txtTool.SelStart = 0
        txtTool.SelLength = Len(txtTool.Text)
    ElseIf txtTool < 10 Then
        If Len(txtTool) = 1 Then
            txtTool = "0" & txtTool
        End If
    End If
End Sub

Private Sub txtToolCorrection_Change()
    If Not IsNumeric(txtToolCorrection.Text) Then
        txtToolCorrection.Text = ""
    End If
End Sub

Private Sub txtToolCorrection_GotFocus()
    txtToolCorrection.SelStart = 0
    txtToolCorrection.SelLength = Len(txtToolCorrection.Text)
End Sub

Private Sub txtToolCorrection_Validate(Cancel As Boolean)
    If txtToolCorrection > 99 Or txtToolCorrection < 1 Then
        Cancel = True
        txtToolCorrection.SelStart = 0
        txtToolCorrection.SelLength = Len(txtToolCorrection.Text)
    ElseIf txtToolCorrection < 10 Then
        If Len(txtToolCorrection) = 1 Then
            txtToolCorrection = "0" & txtToolCorrection
        End If
    End If
End Sub

Private Sub txtCutSpeed_Change()
    If Not IsNumeric(txtCutSpeed.Text) Then
        txtCutSpeed.Text = ""
    End If
End Sub

Private Sub txtCutSpeed_GotFocus()
    txtCutSpeed.SelStart = 0
    txtCutSpeed.SelLength = Len(txtCutSpeed.Text)
End Sub

Private Sub txtCutSpeed_Validate(Cancel As Boolean)
    If Option6.Value = True Then
        If txtCutSpeed > 500 Or txtCutSpeed < 2 Then
            Cancel = True
            txtCutSpeed.SelStart = 0
            txtCutSpeed.SelLength = Len(txtCutSpeed.Text)
        End If
    ElseIf Option7.Value = True Then
        If txtCutSpeed > 5000 Or txtCutSpeed < 10 Then
            Cancel = True
            txtCutSpeed.SelStart = 0
            txtCutSpeed.SelLength = Len(txtCutSpeed.Text)
        End If
    End If
End Sub
