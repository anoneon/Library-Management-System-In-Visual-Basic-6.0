VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form splashForm 
   BorderStyle     =   0  'None
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   4080
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1splash 
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
      MouseIcon       =   "Splash.frx":030A
   End
   Begin LBMSuansu.AutoResize Resize 
      Left            =   5520
      Tag             =   "NO"
      Top             =   120
      _ExtentX        =   714
      _ExtentY        =   714
      AspectRatioValue=   0
   End
   Begin VB.Timer Timer1splash 
      Interval        =   270
      Left            =   6000
      Top             =   120
   End
   Begin VB.Label Label5splash 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label7splash 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label6splash 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   3120
      Width           =   5895
   End
   Begin VB.Label Label4splash 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5760
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label3splash 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2splash 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label1splash 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin LBMSuansu.ucAniGIF ucAniGIF1 
      Height          =   4080
      Left            =   0
      Top             =   0
      Width           =   7080
      _ExtentX        =   15875
      _ExtentY        =   15875
      GIF             =   "Splash.frx":0624
   End
End
Attribute VB_Name = "splashForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()



splashForm.Caption = " Library Management System "


Label1splash.Caption = "LIBRARY"
Label2splash.Caption = "MANAGEMENT"
Label3splash.Caption = "SYSTEM"
Label4splash.Caption = "V 1.0"
Label5splash.Caption = "©2018 Suansu"
Label6splash.Caption = "Loadin Please wait"

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub Timer1splash_Timer()

'If (ProgressBar1splash.Value = 10 Or ProgressBar1splash.Value = 30 Or ProgressBar1splash.Value = 50 Or ProgressBar1splash.Value = 70 Or ProgressBar1splash.Value = 90) Then
'Label6splash.Caption = Label6splash.Caption & "."
'End If

'PBcolor ProgressBar1splash, vbBlack, vbBlue
Label7splash.Caption = ProgressBar1splash.Value & "%"
ProgressBar1splash.Value = ProgressBar1splash.Value + 5

If ProgressBar1splash.Value = ProgressBar1splash.Max Then
Unload Me
login.Show

End If
End Sub

