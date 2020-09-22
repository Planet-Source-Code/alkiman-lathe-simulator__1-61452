VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1095
   ClientLeft      =   2520
   ClientTop       =   6270
   ClientWidth     =   5640
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdZoeken 
      Caption         =   "Find"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   607
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   127
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Find :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   1020
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lStartPos As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim lFound As Long
    lFound = frmNC_Editor.RichTextBox1.Find(Combo1.Text, lStartPos, Len(frmNC_Editor.RichTextBox1.Text))
    lStartPos = lFound + 1
    If lFound > 0 Then
        frmNC_Editor.SetFocus
    Else
        Me.Hide
        MsgBox "Search done. Item not find", vbInformation + vbOKOnly
        Unload Me
    End If
End Sub

Private Sub cmdZoeken_Click()
    Combo1.AddItem (Combo1.Text)
    cmdZoeken.Visible = False
    cmdFind.Visible = True
    cmdFind_Click
End Sub

Private Sub Combo1_Change()
    lStartPos = 1
End Sub

Private Sub Form_Load()
    FormOnTop Me.hWnd, True
    Label1.Top = Label1.Top + 30
    cmdFind.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormOnTop frmNC_Editor.hWnd, True
End Sub
