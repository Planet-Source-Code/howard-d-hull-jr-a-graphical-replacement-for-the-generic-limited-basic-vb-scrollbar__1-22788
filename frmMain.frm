VERSION 5.00
Object = "*\A..\..\..\DYNAMI~1\v1.0\DTVISU~1\dtScrollBar.vbp"
Begin VB.Form frmMain 
   Caption         =   "Dynamic Technologies Visual ScrollBar Demo"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   945
      Index           =   6
      Left            =   4950
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "frmMain.frx":0000
      Top             =   2400
      Width           =   2235
   End
   Begin dtScrollBar.dtVisualScrollBar dtVisualScrollBar 
      Height          =   1485
      Index           =   3
      Left            =   7500
      TabIndex        =   3
      Top             =   1620
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   2619
      ThumbAlignment  =   1
      PictureBackground=   "frmMain.frx":005A
      PictureUp       =   "frmMain.frx":2200
      PictureDown     =   "frmMain.frx":3810
      PictureRight    =   "frmMain.frx":4D4C
      PictureLeft     =   "frmMain.frx":6288
      PictureThumb    =   "frmMain.frx":77C4
      Min             =   -2147483647
      Max             =   -1073741822
      Value           =   -1073741822
      BackColor       =   -2147483643
      LargeChange     =   1
      Enabled         =   0   'False
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   5
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frmMain.frx":8BEE
      Top             =   5130
      Width           =   8475
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "frmMain.frx":8C42
      Top             =   180
      Width           =   8475
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   3
      Left            =   4920
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmMain.frx":8C6F
      Top             =   3330
      Width           =   3795
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   2
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmMain.frx":8D9C
      Top             =   540
      Width           =   8505
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   915
      Index           =   1
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmMain.frx":8ED8
      Top             =   3780
      Width           =   3795
   End
   Begin dtScrollBar.dtVisualScrollBar dtVisualScrollBar 
      Height          =   210
      Index           =   1
      Left            =   180
      TabIndex        =   0
      Top             =   1740
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   370
      ButtonsVisible  =   0   'False
      Orientation     =   0
      ThumbAlignment  =   4
      PictureBackground=   "frmMain.frx":8F3B
      PictureUp       =   "frmMain.frx":9E50
      PictureDown     =   "frmMain.frx":A8D4
      PictureRight    =   "frmMain.frx":B35B
      PictureLeft     =   "frmMain.frx":BDE8
      PictureThumb    =   "frmMain.frx":C878
      Min             =   -2147483647
      Max             =   0
      AutoSize        =   -1  'True
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   915
      Index           =   0
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmMain.frx":D31D
      Top             =   2100
      Width           =   3795
   End
   Begin dtScrollBar.dtVisualScrollBar dtVisualScrollBar 
      Height          =   315
      Index           =   2
      Left            =   180
      TabIndex        =   1
      ToolTipText     =   "Tool Tip text 2"
      Top             =   3360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      Orientation     =   0
      ThumbAlignment  =   4
      PictureBackground=   "frmMain.frx":D37C
      PictureUp       =   "frmMain.frx":D398
      PictureDown     =   "frmMain.frx":DE1C
      PictureRight    =   "frmMain.frx":E8A3
      PictureLeft     =   "frmMain.frx":F330
      PictureThumb    =   "frmMain.frx":FDC0
      Max             =   1
      Value           =   1
      BackColor       =   8421631
   End
   Begin dtScrollBar.dtVisualScrollBar dtVisualScrollBar 
      Height          =   525
      Index           =   0
      Left            =   4920
      TabIndex        =   2
      Top             =   1680
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   926
      Orientation     =   0
      ThumbAlignment  =   4
      PictureBackground=   "frmMain.frx":107DC
      PictureUp       =   "frmMain.frx":10CE6
      PictureDown     =   "frmMain.frx":1176A
      PictureRight    =   "frmMain.frx":121F1
      PictureLeft     =   "frmMain.frx":12492
      PictureThumb    =   "frmMain.frx":12749
      Max             =   2000000000
      Value           =   1
      BackColor       =   255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dtVisualScrollBar_Change(Index As Integer)
   Me.Caption = dtVisualScrollBar(Index).Value
End Sub

Private Sub dtVisualScrollBar_Scroll(Index As Integer)
   Me.Caption = dtVisualScrollBar(Index).Value
End Sub

