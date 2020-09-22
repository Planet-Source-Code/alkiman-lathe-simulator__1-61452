VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmG9094 
   Caption         =   "Automatic Cycle"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   Icon            =   "frmG9094.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Index           =   90
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   9551
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "G90 Outer/Inner Cycle"
      TabPicture(0)   =   "frmG9094.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "HScroll1(90)"
      Tab(0).Control(1)=   "HScroll2(90)"
      Tab(0).Control(2)=   "lstSample"
      Tab(0).Control(3)=   "chkKoelwater(90)"
      Tab(0).Control(4)=   "cmdCancel(90)"
      Tab(0).Control(5)=   "cmdAdd(90)"
      Tab(0).Control(6)=   "picVoorbeeld(90)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdVoorbeeld(90)"
      Tab(0).Control(8)=   "txtstapX(90)"
      Tab(0).Control(9)=   "txteindX(90)"
      Tab(0).Control(10)=   "txtFeed(90)"
      Tab(0).Control(11)=   "txtConisch(90)"
      Tab(0).Control(12)=   "txtstartZ(90)"
      Tab(0).Control(13)=   "txtstartX(90)"
      Tab(0).Control(14)=   "txtsPosZ(90)"
      Tab(0).Control(15)=   "txtsPosX(90)"
      Tab(0).Control(16)=   "lblZoom(90)"
      Tab(0).Control(17)=   "lblZoomFactor(2)"
      Tab(0).Control(18)=   "lblSpeed(90)"
      Tab(0).Control(19)=   "lblDrawSpeed(2)"
      Tab(0).Control(20)=   "Label11(0)"
      Tab(0).Control(21)=   "Label10(0)"
      Tab(0).Control(22)=   "Label9(0)"
      Tab(0).Control(23)=   "Label8(0)"
      Tab(0).Control(24)=   "Label7(0)"
      Tab(0).Control(25)=   "Label6(0)"
      Tab(0).Control(26)=   "Label5(0)"
      Tab(0).Control(27)=   "Label4(0)"
      Tab(0).Control(28)=   "Label3(0)"
      Tab(0).Control(29)=   "Label2(0)"
      Tab(0).Control(30)=   "Label1(0)"
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "G92 Thread Cycle"
      TabPicture(1)   =   "frmG9094.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "HScroll1(92)"
      Tab(1).Control(1)=   "HScroll2(92)"
      Tab(1).Control(2)=   "chkKoelwater(92)"
      Tab(1).Control(3)=   "txtsPosX(92)"
      Tab(1).Control(4)=   "txtsPosZ(92)"
      Tab(1).Control(5)=   "txtstartX(92)"
      Tab(1).Control(6)=   "txtstartZ(92)"
      Tab(1).Control(7)=   "txtConisch(92)"
      Tab(1).Control(8)=   "txtFeed(92)"
      Tab(1).Control(9)=   "txteindX(92)"
      Tab(1).Control(10)=   "txtstapX(92)"
      Tab(1).Control(11)=   "cmdVoorbeeld(92)"
      Tab(1).Control(12)=   "picVoorbeeld(92)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmdAdd(92)"
      Tab(1).Control(14)=   "cmdCancel(92)"
      Tab(1).Control(15)=   "lblZoom(92)"
      Tab(1).Control(16)=   "lblZoomFactor(1)"
      Tab(1).Control(17)=   "lblSpeed(92)"
      Tab(1).Control(18)=   "lblDrawSpeed(1)"
      Tab(1).Control(19)=   "Label1(1)"
      Tab(1).Control(20)=   "Label2(1)"
      Tab(1).Control(21)=   "Label3(1)"
      Tab(1).Control(22)=   "Label4(1)"
      Tab(1).Control(23)=   "Label5(1)"
      Tab(1).Control(24)=   "Label6(1)"
      Tab(1).Control(25)=   "Label7(1)"
      Tab(1).Control(26)=   "Label8(1)"
      Tab(1).Control(27)=   "Label9(1)"
      Tab(1).Control(28)=   "Label10(1)"
      Tab(1).Control(29)=   "Label11(1)"
      Tab(1).ControlCount=   30
      TabCaption(2)   =   "G94 End/Face Cycle"
      TabPicture(2)   =   "frmG9094.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label11(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label10(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label9(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label8(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label7(2)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label6(2)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label5(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label4(2)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label3(2)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label2(2)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label1(2)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "lblZoom(94)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "lblZoomFactor(0)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "lblSpeed(94)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "lblDrawSpeed(3)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "cmdCancel(94)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "cmdAdd(94)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "picVoorbeeld(94)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "cmdVoorbeeld(94)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "txtstapX(94)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "txteindX(94)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "txtFeed(94)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "txtConisch(94)"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "txtstartZ(94)"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "txtstartX(94)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "txtsPosZ(94)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "txtsPosX(94)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "chkKoelwater(94)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "HScroll1(94)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "HScroll2(94)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).ControlCount=   30
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Index           =   94
         LargeChange     =   10
         Left            =   240
         Max             =   0
         Min             =   20
         TabIndex        =   88
         Top             =   4800
         Value           =   10
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Index           =   94
         LargeChange     =   10
         Left            =   240
         Max             =   200
         Min             =   10
         TabIndex        =   87
         Top             =   4200
         Value           =   10
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Index           =   90
         LargeChange     =   10
         Left            =   -74760
         Max             =   200
         Min             =   10
         TabIndex        =   82
         Top             =   4200
         Value           =   10
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Index           =   90
         LargeChange     =   10
         Left            =   -74760
         Max             =   0
         Min             =   20
         TabIndex        =   81
         Top             =   4800
         Value           =   10
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Index           =   92
         LargeChange     =   10
         Left            =   -74760
         Max             =   200
         Min             =   10
         TabIndex        =   76
         Top             =   4200
         Value           =   10
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Index           =   92
         LargeChange     =   10
         Left            =   -74760
         Max             =   0
         Min             =   20
         TabIndex        =   75
         Top             =   4800
         Value           =   10
         Width           =   1335
      End
      Begin VB.ListBox lstSample 
         Height          =   255
         Left            =   -68280
         TabIndex        =   73
         Top             =   1560
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox chkKoelwater 
         Caption         =   "Coolant On"
         Height          =   255
         Index           =   94
         Left            =   5100
         TabIndex        =   27
         Top             =   600
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkKoelwater 
         Caption         =   "Coolant On"
         Height          =   255
         Index           =   92
         Left            =   -69900
         TabIndex        =   15
         Top             =   600
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkKoelwater 
         Caption         =   "Coolant On"
         Height          =   255
         Index           =   90
         Left            =   -69900
         TabIndex        =   3
         Top             =   600
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox txtsPosX 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   94
         Left            =   2220
         TabIndex        =   25
         Text            =   "84"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtsPosZ 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   94
         Left            =   3660
         TabIndex        =   26
         Text            =   "5"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtstartX 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   94
         Left            =   2220
         TabIndex        =   28
         Text            =   "-1.6"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtstartZ 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   94
         Left            =   3660
         TabIndex        =   29
         Text            =   "3"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtConisch 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   94
         Left            =   5100
         TabIndex        =   30
         Text            =   "0"
         ToolTipText     =   "Conisiteit"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtFeed 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   94
         Left            =   6540
         TabIndex        =   31
         Text            =   "0.2"
         ToolTipText     =   "Voeding"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txteindX 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   94
         Left            =   3660
         TabIndex        =   32
         Text            =   "0.1"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtstapX 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   94
         Left            =   5100
         TabIndex        =   33
         Text            =   "1"
         ToolTipText     =   "Stap per snede"
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdVoorbeeld 
         Caption         =   "Example"
         Height          =   375
         Index           =   94
         Left            =   480
         TabIndex        =   34
         Top             =   2160
         Width           =   1215
      End
      Begin VB.PictureBox picVoorbeeld 
         AutoRedraw      =   -1  'True
         Height          =   2895
         Index           =   94
         Left            =   2040
         ScaleHeight     =   189
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   2160
         Width           =   5535
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Index           =   94
         Left            =   480
         TabIndex        =   35
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Close"
         Height          =   375
         Index           =   94
         Left            =   480
         TabIndex        =   36
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtsPosX 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   92
         Left            =   -72780
         TabIndex        =   13
         Text            =   "64"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtsPosZ 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   92
         Left            =   -71340
         TabIndex        =   14
         Text            =   "2"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtstartX 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   92
         Left            =   -72780
         TabIndex        =   16
         Text            =   "60"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtstartZ 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   92
         Left            =   -71340
         TabIndex        =   17
         Text            =   "-40"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtConisch 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   92
         Left            =   -69900
         TabIndex        =   18
         Text            =   "0"
         ToolTipText     =   "Conisiteit"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtFeed 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   92
         Left            =   -68460
         TabIndex        =   19
         Text            =   "1"
         ToolTipText     =   "Voeding"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txteindX 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   92
         Left            =   -72780
         TabIndex        =   20
         Text            =   "58"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtstapX 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   92
         Left            =   -69900
         TabIndex        =   21
         Text            =   "0.2"
         ToolTipText     =   "Stap per snede"
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdVoorbeeld 
         Caption         =   "Example"
         Height          =   375
         Index           =   92
         Left            =   -74520
         TabIndex        =   22
         Top             =   2160
         Width           =   1215
      End
      Begin VB.PictureBox picVoorbeeld 
         AutoRedraw      =   -1  'True
         Height          =   2895
         Index           =   92
         Left            =   -72960
         ScaleHeight     =   189
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2160
         Width           =   5535
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Index           =   92
         Left            =   -74520
         TabIndex        =   23
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Close"
         Height          =   375
         Index           =   92
         Left            =   -74520
         TabIndex        =   24
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Close"
         Height          =   375
         Index           =   90
         Left            =   -74520
         TabIndex        =   12
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Index           =   90
         Left            =   -74520
         TabIndex        =   11
         Top             =   2760
         Width           =   1215
      End
      Begin VB.PictureBox picVoorbeeld 
         AutoRedraw      =   -1  'True
         Height          =   2880
         Index           =   90
         Left            =   -72960
         ScaleHeight     =   188
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2160
         Width           =   5535
      End
      Begin VB.CommandButton cmdVoorbeeld 
         Caption         =   "Example"
         Height          =   375
         Index           =   90
         Left            =   -74520
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtstapX 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   90
         Left            =   -69900
         TabIndex        =   9
         Text            =   "3"
         ToolTipText     =   "Stap per snede"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txteindX 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   90
         Left            =   -72780
         TabIndex        =   8
         Text            =   "60"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtFeed 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   90
         Left            =   -68460
         TabIndex        =   7
         Text            =   "0.2"
         ToolTipText     =   "Voeding"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtConisch 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   90
         Left            =   -69900
         TabIndex        =   6
         Text            =   "0"
         ToolTipText     =   "Conisiteit"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtstartZ 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   90
         Left            =   -71340
         TabIndex        =   5
         Text            =   "-60"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtstartX 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   90
         Left            =   -72780
         TabIndex        =   4
         Text            =   "80"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtsPosZ 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   90
         Left            =   -71340
         TabIndex        =   2
         Text            =   "2"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtsPosX 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   90
         Left            =   -72780
         TabIndex        =   1
         Text            =   "84"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblDrawSpeed 
         Caption         =   "Draw speed"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   92
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label lblSpeed 
         Alignment       =   2  'Center
         Caption         =   "10"
         Height          =   255
         Index           =   94
         Left            =   1560
         TabIndex        =   91
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label lblZoomFactor 
         Caption         =   "Zoom factor"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   90
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lblZoom 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   255
         Index           =   94
         Left            =   1560
         TabIndex        =   89
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label lblZoom 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   255
         Index           =   90
         Left            =   -73440
         TabIndex        =   86
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label lblZoomFactor 
         Caption         =   "Zoom factor"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   85
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lblSpeed 
         Alignment       =   2  'Center
         Caption         =   "10"
         Height          =   255
         Index           =   90
         Left            =   -73440
         TabIndex        =   84
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label lblDrawSpeed 
         Caption         =   "Draw speed"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   83
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label lblZoom 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   255
         Index           =   92
         Left            =   -73440
         TabIndex        =   80
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label lblZoomFactor 
         Caption         =   "Zoom factor"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   79
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lblSpeed 
         Alignment       =   2  'Center
         Caption         =   "10"
         Height          =   255
         Index           =   92
         Left            =   -73440
         TabIndex        =   78
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label lblDrawSpeed 
         Caption         =   "Draw speed"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   77
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label lblDrawSpeed 
         Caption         =   "Draw speed"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   74
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Start position :"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   72
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   1980
         TabIndex        =   71
         Top             =   630
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   3420
         TabIndex        =   70
         Top             =   630
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "First Cut :"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   69
         Top             =   1110
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   1980
         TabIndex        =   68
         Top             =   1110
         Width           =   135
      End
      Begin VB.Label Label6 
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   3420
         TabIndex        =   67
         Top             =   1110
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4860
         TabIndex        =   66
         Top             =   1110
         Width           =   135
      End
      Begin VB.Label Label8 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6300
         TabIndex        =   65
         Top             =   1110
         Width           =   135
      End
      Begin VB.Label Label9 
         Caption         =   "Last Cut :"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   64
         Top             =   1590
         Width           =   1275
      End
      Begin VB.Label Label10 
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   3420
         TabIndex        =   63
         Top             =   1590
         Width           =   135
      End
      Begin VB.Label Label11 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4860
         TabIndex        =   62
         Top             =   1590
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Start position :"
         Height          =   255
         Index           =   1
         Left            =   -74520
         TabIndex        =   60
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   -73020
         TabIndex        =   59
         Top             =   630
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   -71580
         TabIndex        =   58
         Top             =   630
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "First Cut :"
         Height          =   255
         Index           =   1
         Left            =   -74520
         TabIndex        =   57
         Top             =   1110
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   -73020
         TabIndex        =   56
         Top             =   1110
         Width           =   135
      End
      Begin VB.Label Label6 
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   -71580
         TabIndex        =   55
         Top             =   1110
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -70140
         TabIndex        =   54
         Top             =   1110
         Width           =   135
      End
      Begin VB.Label Label8 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -68700
         TabIndex        =   53
         Top             =   1110
         Width           =   135
      End
      Begin VB.Label Label9 
         Caption         =   "Last Cut :"
         Height          =   255
         Index           =   1
         Left            =   -74520
         TabIndex        =   52
         Top             =   1590
         Width           =   1275
      End
      Begin VB.Label Label10 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   -73020
         TabIndex        =   51
         Top             =   1590
         Width           =   135
      End
      Begin VB.Label Label11 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -70140
         TabIndex        =   50
         Top             =   1590
         Width           =   135
      End
      Begin VB.Label Label11 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -70140
         TabIndex        =   47
         Top             =   1590
         Width           =   135
      End
      Begin VB.Label Label10 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   -73020
         TabIndex        =   46
         Top             =   1590
         Width           =   135
      End
      Begin VB.Label Label9 
         Caption         =   "Last Cut :"
         Height          =   255
         Index           =   0
         Left            =   -74520
         TabIndex        =   45
         Top             =   1590
         Width           =   1275
      End
      Begin VB.Label Label8 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -68700
         TabIndex        =   44
         Top             =   1110
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -70140
         TabIndex        =   43
         Top             =   1110
         Width           =   135
      End
      Begin VB.Label Label6 
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   -71580
         TabIndex        =   42
         Top             =   1110
         Width           =   135
      End
      Begin VB.Label Label5 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   -73020
         TabIndex        =   41
         Top             =   1110
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "First Cut :"
         Height          =   255
         Index           =   0
         Left            =   -74520
         TabIndex        =   40
         Top             =   1110
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   -71580
         TabIndex        =   39
         Top             =   630
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   -73020
         TabIndex        =   38
         Top             =   630
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Start position :"
         Height          =   255
         Index           =   0
         Left            =   -74520
         TabIndex        =   37
         Top             =   630
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmG9094"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click(Index As Integer)
    Dim strTool As String
    Dim dStep As Double
    strTool = "G00 X" & txtsPosX(Index) & " Z" & txtsPosZ(Index) & IIf(chkKoelwater(Index).Value = 1, " M08 ;", " ;") & vbCrLf
    If txtConisch(Index) = 0 Then
        strTool = strTool & "G" & Index & " X" & txtstartX(Index) & " Z" & txtstartZ(Index) & " F" & txtFeed(Index) & " ;" & vbCrLf
    Else
        strTool = strTool & "G" & Index & " X" & txtstartX(Index) & " Z" & txtstartZ(Index) & " R" & txtConisch(Index) & " F" & txtFeed(Index) & " ;" & vbCrLf
    End If
    If Index = 94 Then
        dStep = Val(txtstartZ(Index).Text) - Val(txtstapX(Index).Text)
    Else
        If txtstartX(Index).Text > txteindX(Index).Text Then
            dStep = Val(txtstartX(Index).Text) - Val(txtstapX(Index).Text)
        Else
            dStep = Val(txtstartX(Index).Text) + Val(txtstapX(Index).Text)
        End If
    End If
    If IIf(Index = 94, txtstartZ(Index).Text > txteindX(Index).Text, txtstartX(Index).Text > txteindX(Index).Text) Then
        Debug.Print Val(txteindX(Index).Text)
        Do Until dStep < Val(txteindX(Index).Text)
            If Index = 94 Then
                strTool = strTool & "Z" & Trim(Str(dStep)) & " ;" & vbCrLf
            Else
                strTool = strTool & "X" & Trim(Str(dStep)) & " ;" & vbCrLf
            End If
            dStep = dStep - Val(txtstapX(Index).Text)
        Loop
    Else
        Do Until dStep > Val(txteindX(Index).Text)
            If Index = 94 Then
                strTool = strTool & "Z" & Trim(Str(dStep)) & " ;" & vbCrLf
            Else
                strTool = strTool & "X" & Trim(Str(dStep)) & " ;" & vbCrLf
            End If
            dStep = dStep + Val(txtstapX(Index).Text)
        Loop
    End If
    If Val(dStep + txtstapX(Index).Text) > Val(txteindX(Index).Text) Then
        If Index = 94 Then
            strTool = strTool & "Z" & txteindX(Index).Text & " ;" & vbCrLf
        Else
            strTool = strTool & "X" & txteindX(Index).Text & " ;" & vbCrLf
        End If
    End If
    strTool = strTool & "G00 X" & txtsPosX(Index).Text & " Z" & txtsPosZ(Index).Text & " ;" & vbCrLf
    Clipboard.Clear
    Clipboard.SetText strTool, vbCFText
    frmGraphic.RichTextBox1.SelText = Clipboard.GetText(vbCFText)
    'Unload Me
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdVoorbeeld_Click(Index As Integer)
    Debug.Print Index
    Dim strTool As String
    Dim dStep As Double
    lstSample.Clear
    lstSample.AddItem "G00 X" & txtsPosX(Index) & " Z" & txtsPosZ(Index) & IIf(chkKoelwater(Index).Value = 1, " M08 ;", " ;")
    If txtConisch(Index) = 0 Then
        lstSample.AddItem "G" & Index & " X" & txtstartX(Index) & " Z" & txtstartZ(Index) & " F" & txtFeed(Index) & " ;"
    Else
        lstSample.AddItem "G" & Index & " X" & txtstartX(Index) & " Z" & txtstartZ(Index) & " R" & txtConisch(Index) & " F" & txtFeed(Index) & " ;"
    End If
    If Index = 94 Then
        dStep = Val(txtstartZ(Index).Text) - Val(txtstapX(Index).Text)
    Else
        If txtstartX(Index).Text > txteindX(Index).Text Then
            dStep = Val(txtstartX(Index).Text) - Val(txtstapX(Index).Text)
        Else
            dStep = Val(txtstartX(Index).Text) + Val(txtstapX(Index).Text)
        End If
    End If
    If IIf(Index = 94, txtstartZ(Index).Text > txteindX(Index).Text, txtstartX(Index).Text > txteindX(Index).Text) Then
        Debug.Print Val(txteindX(Index).Text)
        Do Until dStep < Val(txteindX(Index).Text)
            If Index = 94 Then
                lstSample.AddItem "Z" & Trim(Str(dStep)) & " ;"
            Else
                lstSample.AddItem "X" & Trim(Str(dStep)) & " ;"
            End If
            dStep = dStep - Val(txtstapX(Index).Text)
        Loop
    Else
        Do Until dStep > Val(txteindX(Index).Text)
            If Index = 94 Then
                lstSample.AddItem "Z" & Trim(Str(dStep)) & " ;"
            Else
                lstSample.AddItem "X" & Trim(Str(dStep)) & " ;"
            End If
            dStep = dStep + Val(txtstapX(Index).Text)
        Loop
    End If
    If Val(dStep + txtstapX(Index).Text) > Val(txteindX(Index).Text) Then
        If Index = 94 Then
            lstSample.AddItem "Z" & txteindX(Index).Text & " ;"
        Else
            lstSample.AddItem "X" & txteindX(Index).Text & " ;"
        End If
    End If
    lstSample.AddItem "G00 X" & txtsPosX(Index).Text & " Z" & txtsPosZ(Index).Text & " ;"
    Dim sData As String
    Dim sLockPoint As String
    Dim dsXpos As Double, dsYpos As Double, dExtraX As Double
    xWidth = 330
    yHeight = 170
    picVoorbeeld(Index).Cls
    picVoorbeeld(Index).DrawStyle = 0
    picVoorbeeld(Index).ForeColor = vbBlue
    If Index <> 94 Then
        dsYpos = yHeight - (Val(txtstartX(Index).Text) * dScaleFactor / 2)
        picVoorbeeld(Index).Line (10, dsYpos)-((xWidth + 30), dsYpos)
        dsYpos = yHeight - (Val(txteindX(Index).Text) * dScaleFactor / 2)
        picVoorbeeld(Index).Line (10, dsYpos)-((xWidth + 30), dsYpos)
    Else
        dsYpos = xWidth + (Val(txteindX(Index).Text) * dScaleFactor)
        picVoorbeeld(Index).Line (dsYpos, 10)-(dsYpos, (yHeight + 20))
        dsYpos = xWidth + (Val(txtstartZ(Index).Text) * dScaleFactor)
        picVoorbeeld(Index).Line (dsYpos, 10)-(dsYpos, (yHeight + 20))
    End If
    lLineColor = &HFF00FF
    picVoorbeeld(Index).ForeColor = lLineColor
    picVoorbeeld(Index).DrawStyle = 2 'ijlgang
    sData = "G00 X250 Z250 M09 ;"
    DrawContour sData, False, picVoorbeeld(Index), False
    If Index = 94 Then
        dsYpos = txtsPosX(Index)
        dsXpos = txtsPosZ(Index)
        sLockPoint = "G01 X" & txtstartX(Index)
        dExtraX = 0
        If txtConisch(Index) <> 0 Then
            dExtraX = txtConisch(Index)
        End If
        lstSample.TopIndex = 0
        sData = lstSample.List(0)
        DrawContour sData, False, picVoorbeeld(Index), False
        lstSample.TopIndex = 1
        sData = lstSample.List(1)
        sData = "G00 Z" & GetCoordinates(sData, InStr(1, sData, "Z"), False) + dExtraX & " ;"
        DrawContour sData, False, picVoorbeeld(Index), False
        picVoorbeeld(Index).DrawStyle = 0
        sData = sLockPoint & " Z" & GetCoordinates(sData, InStr(1, sData, "Z"), False) - dExtraX & " ;"
        DrawContour sData, False, picVoorbeeld(Index), False
        sData = "Z" & dsXpos & " ;"
        DrawContour sData, False, picVoorbeeld(Index), False
        picVoorbeeld(Index).DrawStyle = 2 'ijlgang
        sData = "G00 X" & dsYpos & " ;"
        DrawContour sData, False, picVoorbeeld(Index), False
        lstSample.ListIndex = 1
        Do
            lstSample.ListIndex = lstSample.ListIndex + 1
            lstSample.TopIndex = lstSample.ListIndex
            sData = lstSample.List(lstSample.ListIndex)
            If InStr(1, sData, "G") = 0 Then
                sData = "Z" & GetCoordinates(sData, InStr(1, sData, "Z"), False) + dExtraX & " ;"
                DrawContour sData, False, picVoorbeeld(Index), False
                picVoorbeeld(Index).DrawStyle = 0
                sData = sLockPoint & " Z" & GetCoordinates(sData, InStr(1, sData, "Z"), False) - dExtraX & " ;"
                DrawContour sData, False, picVoorbeeld(Index), False
                sData = "Z" & dsXpos & " ;"
                DrawContour sData, False, picVoorbeeld(Index), False
                picVoorbeeld(Index).DrawStyle = 2 'ijlgang
                sData = "G00 X" & dsYpos & " ;"
                DrawContour sData, False, picVoorbeeld(Index), False
                sData = lstSample.List(lstSample.ListIndex)
            End If
        Loop Until InStr(1, sData, "G") <> 0
    Else
        dsYpos = txtsPosX(Index)
        dsXpos = txtsPosZ(Index)
        sLockPoint = "G01 Z" & txtstartZ(Index)
        dExtraX = 0
        If txtConisch(Index) <> 0 Then
            dExtraX = txtConisch(Index) * 2
        End If
        lstSample.TopIndex = 0
        picVoorbeeld(Index).DrawStyle = 2 'ijlgang
        sData = lstSample.List(0)
        DrawContour sData, False, picVoorbeeld(Index), False
        lstSample.TopIndex = 1
        sData = lstSample.List(1)
        sData = "G00 X" & GetCoordinates(sData, InStr(1, sData, "X"), False) + dExtraX & " ;"
        DrawContour sData, False, picVoorbeeld(Index), False
        picVoorbeeld(Index).DrawStyle = 0
        sData = sLockPoint & " X" & GetCoordinates(sData, InStr(1, sData, "X"), False) - dExtraX & " ;"
        DrawContour sData, False, picVoorbeeld(Index), False
        sData = "X" & dsYpos & " ;"
        DrawContour sData, False, picVoorbeeld(Index), False
        picVoorbeeld(Index).DrawStyle = 2 'ijlgang
        sData = "G00 Z" & dsXpos & " ;"
        DrawContour sData, False, picVoorbeeld(Index), False
        lstSample.ListIndex = 1
        Do
            lstSample.ListIndex = lstSample.ListIndex + 1
            lstSample.TopIndex = lstSample.ListIndex
            sData = lstSample.List(lstSample.ListIndex)
            If InStr(1, sData, "G") = 0 Then
                sData = "X" & GetCoordinates(sData, InStr(1, sData, "X"), False) + dExtraX & " ;"
                DrawContour sData, False, picVoorbeeld(Index), False
                picVoorbeeld(Index).DrawStyle = 0
                sData = sLockPoint & " X" & GetCoordinates(sData, InStr(1, sData, "X"), False) - dExtraX & " ;"
                DrawContour sData, False, picVoorbeeld(Index), False
                sData = "X" & dsYpos & " ;"
                DrawContour sData, False, picVoorbeeld(Index), False
                picVoorbeeld(Index).DrawStyle = 2 'ijlgang
                sData = "G00 Z" & dsXpos & " ;"
                DrawContour sData, False, picVoorbeeld(Index), False
                sData = lstSample.List(lstSample.ListIndex)
            End If
        Loop Until InStr(1, sData, "G") <> 0
    End If
    sData = "G00 X250 Z250 M09 ;"
    DrawContour sData, False, picVoorbeeld(Index), False
    xWidth = 660
    yHeight = 340
End Sub

Private Sub Form_Load()
    Me.Top = frmGraphic.Top + 945
    Me.Left = frmGraphic.Left + 2415
    HScroll1(90).Value = dScaleFactor * 10
    HScroll1(92).Value = dScaleFactor * 10
    HScroll1(94).Value = dScaleFactor * 10
    HScroll2(90).Value = lTimeMilliSeconds
    HScroll2(92).Value = lTimeMilliSeconds
    HScroll2(94).Value = lTimeMilliSeconds
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmGraphic.HScroll1.Value = dScaleFactor * 10
    frmGraphic.HScroll2.Value = lTimeMilliSeconds
End Sub

Private Sub HScroll1_Change(Index As Integer)
    lblZoom(Index).Caption = HScroll1(Index).Value / 10
    dScaleFactor = HScroll1(Index).Value / 10
End Sub

Private Sub HScroll1_LostFocus(Index As Integer)
    Dim i As Integer
    For i = 90 To 94 Step 2
        HScroll1(i).Value = HScroll1(Index).Value
    Next
End Sub

Private Sub HScroll1_Scroll(Index As Integer)
    lblZoom(Index).Caption = HScroll1(Index).Value / 10
    dScaleFactor = HScroll1(Index).Value / 10
End Sub

Private Sub HScroll2_Change(Index As Integer)
    lblSpeed(Index).Caption = (20 - HScroll2(Index).Value) * 5
    lTimeMilliSeconds = HScroll2(Index).Value
End Sub

Private Sub HScroll2_LostFocus(Index As Integer)
    Dim i As Integer
    For i = 90 To 94 Step 2
        HScroll2(i).Value = HScroll2(Index).Value
    Next
End Sub

Private Sub HScroll2_Scroll(Index As Integer)
    lblSpeed(Index).Caption = (20 - HScroll2(Index).Value) * 5
    lTimeMilliSeconds = HScroll2(Index).Value
End Sub

Private Sub txtConisch_Change(Index As Integer)
    If Left(txtConisch(Index), 1) = "-" Then
        If Len(txtConisch(Index)) > 1 Then
            If Not IsNumeric(Mid(txtConisch(Index).Text, 2)) Then
                txtConisch(Index).Text = ""
            End If
        End If
    Else
        If Not IsNumeric(txtConisch(Index).Text) Then
            txtConisch(Index).Text = ""
        End If
    End If
End Sub

Private Sub txtConisch_GotFocus(Index As Integer)
    txtConisch(Index).SelStart = 0
    txtConisch(Index).SelLength = Len(txtConisch(Index).Text)
End Sub

Private Sub txtConisch_Validate(Index As Integer, Cancel As Boolean)
    If txtConisch(Index) = "" Or txtConisch(Index) = "-" Then
        Cancel = True
        txtConisch(Index).SelStart = 0
        txtConisch(Index).SelLength = Len(txtConisch(Index).Text)
    End If
End Sub

Private Sub txteindX_Change(Index As Integer)
    If Index = 94 Then
        If Left(txteindX(Index), 1) = "-" Then
            If Len(txteindX(Index)) > 1 Then
                If Not IsNumeric(Mid(txteindX(Index).Text, 2)) Then
                    txteindX(Index).Text = ""
                End If
            End If
        Else
            If Not IsNumeric(txteindX(Index).Text) Then
                txteindX(Index).Text = ""
            End If
        End If
    Else
        If Not IsNumeric(txteindX(Index).Text) Then
            txteindX(Index).Text = ""
        End If
    End If
End Sub

Private Sub txteindX_GotFocus(Index As Integer)
    txteindX(Index).SelStart = 0
    txteindX(Index).SelLength = Len(txteindX(Index).Text)
End Sub

Private Sub txteindX_Validate(Index As Integer, Cancel As Boolean)
    If Index = 94 Then
        If txteindX(Index) = "" Or txteindX(Index) = "-" Then
            Cancel = True
            txteindX(Index).SelStart = 0
            txteindX(Index).SelLength = Len(txteindX(Index).Text)
        ElseIf Val(txteindX(Index)) >= Val(txtstartZ(Index)) Then
            Cancel = True
            txteindX(Index).SelStart = 0
            txteindX(Index).SelLength = Len(txteindX(Index).Text)
        End If
    Else
        If txteindX(Index) = "" Or Val(txteindX(Index)) < 0 Then
            Cancel = True
            txteindX(Index).SelStart = 0
            txteindX(Index).SelLength = Len(txteindX(Index).Text)
        End If
    End If
End Sub

Private Sub txtFeed_Change(Index As Integer)
    If Not IsNumeric(txtFeed(Index).Text) Then
        txtFeed(Index).Text = ""
    End If
End Sub

Private Sub txtFeed_GotFocus(Index As Integer)
    txtFeed(Index).SelStart = 0
    txtFeed(Index).SelLength = Len(txtFeed(Index).Text)
End Sub

Private Sub txtFeed_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtFeed(Index)) <= 0 Then
        Cancel = True
        txtFeed(Index).SelStart = 0
        txtFeed(Index).SelLength = Len(txtFeed(Index).Text)
    End If
End Sub

Private Sub txtsPosX_Change(Index As Integer)
    If Not IsNumeric(txtsPosX(Index).Text) Then
        txtsPosX(Index).Text = ""
    End If
End Sub

Private Sub txtsPosX_GotFocus(Index As Integer)
    txtsPosX(Index).SelStart = 0
    txtsPosX(Index).SelLength = Len(txtsPosX(Index).Text)
End Sub

Private Sub txtsPosX_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtsPosX(Index)) <= 0 Then
        Cancel = True
        txtsPosX(Index).SelStart = 0
        txtsPosX(Index).SelLength = Len(txtsPosX(Index).Text)
    End If
End Sub

Private Sub txtsPosZ_Change(Index As Integer)
    If Left(txtsPosZ(Index), 1) = "-" Then
        If Len(txtsPosZ(Index)) > 1 Then
            If Not IsNumeric(Mid(txtsPosZ(Index).Text, 2)) Then
                txtsPosZ(Index).Text = ""
            End If
        End If
    Else
        If Not IsNumeric(txtsPosZ(Index).Text) Then
            txtsPosZ(Index).Text = ""
        End If
    End If
End Sub

Private Sub txtsPosZ_GotFocus(Index As Integer)
    txtsPosZ(Index).SelStart = 0
    txtsPosZ(Index).SelLength = Len(txtsPosZ(Index).Text)
End Sub

Private Sub txtsPosZ_Validate(Index As Integer, Cancel As Boolean)
    If txtsPosZ(Index) = "" Or txtsPosZ(Index) = "-" Then
        Cancel = True
        txtsPosZ(Index).SelStart = 0
        txtsPosZ(Index).SelLength = Len(txtsPosZ(Index).Text)
    End If
End Sub

Private Sub txtstapX_Change(Index As Integer)
    If Not IsNumeric(txtstapX(Index).Text) Then
        txtstapX(Index).Text = ""
    End If
End Sub

Private Sub txtstapX_GotFocus(Index As Integer)
    txtstapX(Index).SelStart = 0
    txtstapX(Index).SelLength = Len(txtstapX(Index).Text)
End Sub

Private Sub txtstapX_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtstapX(Index)) <= 0 Then
        Cancel = True
        txtstapX(Index).SelStart = 0
        txtstapX(Index).SelLength = Len(txtstapX(Index).Text)
    End If
End Sub

Private Sub txtstartX_Change(Index As Integer)
    If Left(txtstartX(Index), 1) = "-" And Index = 94 Then
        If Len(txtstartX(Index)) > 1 Then
            If Not IsNumeric(Mid(txtstartX(Index).Text, 2)) Then
                txtstartX(Index).Text = ""
            End If
        End If
    ElseIf Not IsNumeric(txtstartX(Index).Text) Then
        txtstartX(Index).Text = ""
    End If
End Sub

Private Sub txtstartX_GotFocus(Index As Integer)
    txtstartX(Index).SelStart = 0
    txtstartX(Index).SelLength = Len(txtstartX(Index).Text)
End Sub

Private Sub txtstartX_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtstartX(Index)) > Val(txtsPosX(Index)) And Index = 94 Then
        Cancel = True
        txtstartX(Index).SelStart = 0
        txtstartX(Index).SelLength = Len(txtstartX(Index).Text)
    ElseIf Val(txtstartX(Index)) <= 0 And Index <> 94 Then
        Cancel = True
        txtstartX(Index).SelStart = 0
        txtstartX(Index).SelLength = Len(txtstartX(Index).Text)
    End If
End Sub

Private Sub txtstartZ_Change(Index As Integer)
    If Left(txtstartZ(Index), 1) = "-" Then
        If Len(txtstartZ(Index)) > 1 Then
            If Not IsNumeric(Mid(txtstartZ(Index).Text, 2)) Then
                txtstartZ(Index).Text = ""
            End If
        End If
    Else
        If Not IsNumeric(txtstartZ(Index).Text) Then
            txtstartZ(Index).Text = ""
        End If
    End If
End Sub

Private Sub txtstartZ_GotFocus(Index As Integer)
    txtstartZ(Index).SelStart = 0
    txtstartZ(Index).SelLength = Len(txtstartZ(Index).Text)
End Sub

Private Sub txtstartZ_LostFocus(Index As Integer)
    If Val(txtstartZ(90).Text) <> Val(txtstartZ(92).Text) And Val(txtConisch(90).Text) <> 0 Then
        Dim zDiff As Double
        zDiff = Val(txtstartZ(92).Text) / Val(txtstartZ(90).Text)
        txtConisch(92).Text = txtConisch(90).Text * zDiff
        txtstartX(92).Text = Val(txteindX(90).Text) + (Val(txtConisch(92).Text))
        If Val(txtsPosX(90).Text) > Val(txtstartX(90).Text) Then
            txteindX(92).Text = Val(txtstartX(92).Text) - Abs(Val(txtConisch(92).Text))
        Else
            txteindX(92).Text = Val(txtstartX(92).Text) + Abs(Val(txtConisch(92).Text))
        End If
    End If
End Sub

Private Sub txtstartZ_Validate(Index As Integer, Cancel As Boolean)
    If txtstartZ(Index) = "" Or txtstartZ(Index) = "-" Then
        Cancel = True
        txtstartZ(Index).SelStart = 0
        txtstartZ(Index).SelLength = Len(txtstartZ(Index).Text)
    ElseIf Val(txtstartZ(Index)) >= Val(txtsPosZ(Index)) Then
        Cancel = True
        txtstartZ(Index).SelStart = 0
        txtstartZ(Index).SelLength = Len(txtstartZ(Index).Text)
    End If
End Sub
