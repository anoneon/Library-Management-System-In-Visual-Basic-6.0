VERSION 5.00
Begin VB.Form login 
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6615
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin Project1.AutoResize AutoResize1 
      Left            =   6000
      Tag             =   "NO"
      Top             =   120
      _ExtentX        =   714
      _ExtentY        =   714
      AspectRatioValue=   0
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   2760
      Top             =   240
   End
   Begin VB.CommandButton Command1login 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text2login 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Text1login 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label4login 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label3login 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label LabelT 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Time :"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label LabelDt 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Date : "
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2login 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1login 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   960
      X2              =   960
      Y1              =   0
      Y2              =   3960
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   855
      Index           =   3
      Left            =   960
      Top             =   1920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   855
      Index           =   2
      Left            =   0
      Top             =   1920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   855
      Index           =   1
      Left            =   960
      Top             =   1080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   855
      Index           =   0
      Left            =   0
      Top             =   1080
      Width           =   975
   End
   Begin VB.Image Image1login 
      Height          =   3855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cnt As Integer

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command1login_Click()
    
    If Text1login.Text = "" Or Text2login.Text = "" Then
        MsgBox "Login Credential is Empty"
        Exit Sub
        
    End If
    
If Left(Text1login.Text, 1) = "'" And Left(Text2login.Text, 1) = "'" Then
            MsgBox ("Dont Try Injection")
            Unload Me
Else: Set rs = cn.Execute("select * from Users where username='" & Text1login.Text & "' and password='" & Text2login.Text & "'")
    If rs.EOF Then
        If cnt < 2 Then
            cnt = cnt + 1
            If Text1login.Text = "" Or Text2login.Text = "" Then
                MsgBox "Login Credentials Can't Be Empty"
                Text1login.SetFocus
            Else
                
                MsgBox ("The Username and Password dont Match")
                Text1login.Text = ""
                Text1login.SetFocus
                Text2login.Text = ""
            End If
            
        
        Else
             MsgBox ("Maximum attempt achieved with wrong data")
             Command1login.Enabled = False
             Text1login.Enabled = False
             Text1login.Text = ""
             Text2login.Enabled = False
             Text2login.Text = ""
             Command1.Visible = True
        
        End If
    Else
        Unload Me
        mainWindow.Show
    End If
End If
End Sub


Private Sub Form_Load()

    
    Me.Icon = LoadPicture(App.Path & "\images\lbm_ico.ico")
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    login.Caption = "Library Management System V 1.0"
    Image1login.Picture = LoadPicture(App.Path & "\images\login.jpg")
    Shape1(0).BackColor = RGB(224, 108, 52)
    Shape1(1).BackColor = RGB(210, 241, 35)
    Shape1(2).BackColor = RGB(237, 164, 239)
    Shape1(3).BackColor = RGB(168, 63, 214)
    Label1login.Caption = "Username : "
    Label2login.Caption = "Password : "
    Label1login.FontSize = "16"
    Label2login.FontSize = "16"
    Text1login.Text = ""
    Text2login.Text = ""
    Text1login.FontSize = "14"
    Text2login.FontSize = "14"
    Text1login.FontBold = "2"
    Text2login.FontBold = "2"
    Command1login.Caption = "Login"
    
    Command1login.BackColor = RGB(239, 111, 16)
    
    Command1.Visible = False
    Set cn = New ADODB.Connection
    cn.Provider = "microsoft.jet.OLEDB.4.0"
    cn.Open App.Path & "\dbase\dBase.mdb"

    
End Sub






Private Sub Timer1_Timer()

    Label3login.Caption = Date
    Label4login.Caption = Time()
End Sub
