VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form mainWindow 
   Caption         =   "LIBRARY MANAGEMENT"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   11355
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   11355
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   690
      Left            =   9600
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   690
      Left            =   7800
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   7560
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1296
      _Version        =   327682
      Appearance      =   1
      Enabled         =   0   'False
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10680
      Top             =   2640
   End
   Begin Project1.AutoResize Resize 
      Left            =   10680
      Tag             =   "NO"
      Top             =   1800
      _ExtentX        =   714
      _ExtentY        =   714
      AspectRatioValue=   0
   End
   Begin VB.Label Label2 
      Caption         =   "Unavailable"
      Height          =   615
      Left            =   9120
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1main 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "   COLLEGE LIBRARY                                           MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   4680
      Width           =   4455
   End
   Begin VB.Image Image1main 
      Height          =   7215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11415
   End
   Begin VB.Menu bk 
      Caption         =   "&Books"
      Begin VB.Menu adBook 
         Caption         =   "Add Book"
      End
      Begin VB.Menu lsBook 
         Caption         =   "List All Books"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu std 
      Caption         =   "&Student"
      Begin VB.Menu rgStd 
         Caption         =   "Register Student"
      End
      Begin VB.Menu lsStd 
         Caption         =   "List Student"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu brw 
      Caption         =   "&Borrow"
      Begin VB.Menu bkBrw 
         Caption         =   "Borrow Book"
      End
      Begin VB.Menu bkRtr 
         Caption         =   "Return Book"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
   End
   Begin VB.Menu lgt 
      Caption         =   "&Logout"
      Begin VB.Menu lgtt 
         Caption         =   "Log-Out"
      End
      Begin VB.Menu ext 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Abt 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "mainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim ct As Integer

Sub chu()
    
   MsgBox ("See You Again...")
   Timer1.Enabled = True
    
End Sub
Private Sub adBook_Click()
mainWindow.Hide
addBook.Show
End Sub

Private Sub bkBrw_Click()
    Unload Me
    borrowBk.Show
End Sub

Private Sub bkRtr_Click()
    Unload Me
    returnBk.Show
End Sub

Private Sub ext_Click()
Call chu
End Sub

Private Sub Form_Load()
    
    Dim ctt As Integer
    ctt = 0
    
    Me.Icon = LoadPicture(App.Path & "\images\lbm_ico.ico")
    Image1main.Picture = LoadPicture(App.Path & "\images\mainbg.jpg")
    
    Set cn = New ADODB.Connection
    cn.Provider = "microsoft.jet.OLEDB.4.0"
    cn.Open App.Path & "\dbase\dBase.mdb"
    
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "select * from Books ", cn, adOpenDynamic, adLockPessimistic
    
    If Not rs.EOF Then
     If rs.Fields(7).Value > 0 Then
       rs.Fields(8).Value = "YES"
       rs.MoveNext
     Else
       rs.Fields(8).Value = "NO"
       ctt = ctt + 1
       rs.MoveNext
     End If
    End If
    
    Text1.Text = rs.RecordCount
    Text2.Text = ctt

End Sub

Private Sub lgtt_Click()
    Unload Me
    login.Show
End Sub

Private Sub lsBook_Click()
    Unload Me
    listBook.Show
End Sub

Private Sub lsStd_Click()
    Unload Me
    listStd.Show
End Sub

Private Sub rgStd_Click()
    
    Unload Me
    addStd.Show
End Sub

Private Sub Timer1_Timer()

ProgressBar1.Enabled = True

ProgressBar1.Value = ProgressBar1.Value + 10
If ProgressBar1.Value = ProgressBar1.Max Then
Unload Me
End If
End Sub
