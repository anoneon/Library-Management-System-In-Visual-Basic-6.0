VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form mainWindow 
   Caption         =   "LIBRARY MANAGEMENT V 1.0"
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
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   690
      Left            =   9960
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   690
      Left            =   8040
      TabIndex        =   4
      Top             =   720
      Width           =   735
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
   Begin VB.Image Image2 
      Height          =   255
      Index           =   9
      Left            =   3000
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   8
      Left            =   3120
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   7
      Left            =   2160
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   6
      Left            =   840
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   5
      Left            =   2640
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   4
      Left            =   1440
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   3
      Left            =   600
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   2
      Left            =   2520
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   1
      Left            =   1200
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   0
      Left            =   360
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   7215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Unavailable"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Begin VB.Menu opc 
         Caption         =   "OpenIn&Chrome"
      End
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
Option Explicit

Private Declare Function GetMenu Lib "user32" _
   (ByVal hwnd As Long) As Long

Private Declare Function GetSubMenu Lib "user32" _
   (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function SetMenuItemBitmaps Lib "user32" _
   (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, _
    ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Const MF_BYPOSITION = &H400&



Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
     ByVal hwnd As Long, _
     ByVal lpOperation As String, _
     ByVal lpFile As String, _
     ByVal lpParameters As String, _
     ByVal lpDirectory As String, _
     ByVal nShowCmd As Long) As Long

Public Sub OpenChrome(ByVal pURL As String)
    Dim sChromePath As String
    Dim sTmp As String
    Dim sProgramFiles As String
    Dim bNotFound As Boolean
    '
    ' check for 32/64 bit version
    '
    sProgramFiles = Environ("ProgramFiles")
    sChromePath = sProgramFiles & "\Google\Chrome\Application\chrome.exe"
    If Dir$(sChromePath) = vbNullString Then
        ' if not found, search for 32bit version
        sProgramFiles = Environ("ProgramFiles(x86)")
        If sProgramFiles > vbNullString Then
            sChromePath = sProgramFiles & "\Google\Chrome\Application\chrome.exe"
            If Dir$(sChromePath) = vbNullString Then
                bNotFound = True
            End If
        Else
            bNotFound = True
        End If
    End If
    If bNotFound = True Then
        MsgBox "Chrome.exe not found"
        Exit Sub
    End If
    ShellExecute 0, "open", sChromePath, pURL, vbNullString, 1

End Sub







Private Sub SetMenuIcon()
On Error Resume Next
Dim hMenu As Long
Dim hSubMenu As Long
Dim Ret As Long
     'Get main menu ID
     hMenu = GetMenu(hwnd)
     
     
     '
     'MENU FILE
     '
     'Get sub menu 0 (Books items)
     hSubMenu = GetSubMenu(hMenu, 0)
     
     'set bitmap to menu item, by ordinal
     Ret = SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, Image2(0).Picture, Image2(0).Picture)
     Ret = SetMenuItemBitmaps(hSubMenu, 1, MF_BYPOSITION, Image2(1).Picture, Image2(1).Picture)
     
     '
     ' MENU sTUDENT
     '
     'Get sub menu 1 (Stduent items)
     hSubMenu = GetSubMenu(hMenu, 1)
     Ret = SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, Image2(2).Picture, Image2(2).Picture)
     Ret = SetMenuItemBitmaps(hSubMenu, 1, MF_BYPOSITION, Image2(3).Picture, Image2(3).Picture)
     
     
     '
     ' MENU Borrow
     '
     'Get sub menu 3 (Borrow items)
     hSubMenu = GetSubMenu(hMenu, 2)
     Ret = SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, Image2(4).Picture, Image2(4).Picture)
     Ret = SetMenuItemBitmaps(hSubMenu, 1, MF_BYPOSITION, Image2(5).Picture, Image2(5).Picture)
     
     '
     ' MENU Logout
     '
     'Get sub menu 4 (Logout items)
     hSubMenu = GetSubMenu(hMenu, 3)
     Ret = SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, Image2(6).Picture, Image2(6).Picture)
     Ret = SetMenuItemBitmaps(hSubMenu, 1, MF_BYPOSITION, Image2(7).Picture, Image2(7).Picture)
     
      '
     ' MENU Abt
     '
     'Get sub menu 5 (Abt items)
     hSubMenu = GetSubMenu(hMenu, 4)
     Ret = SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, Image2(8).Picture, Image2(8).Picture)
    
     
End Sub

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
    
    Me.Icon = LoadPicture(App.path & "\images\lbm_ico.ico")
    Image1main.Picture = LoadPicture(App.path & "\images\mainbg.jpg")
    
    Set cn = New ADODB.Connection
    cn.Provider = "microsoft.jet.OLEDB.4.0"
    cn.Open App.path & "\dbase\dBase.mdb"
    
    
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
    Image2(0).Picture = LoadPicture(App.path & "\images\addbk.jpg")
    Image2(1).Picture = LoadPicture(App.path & "\images\list.jpg")
    Image2(2).Picture = LoadPicture(App.path & "\images\addstd.jpg")
    Image2(3).Picture = LoadPicture(App.path & "\images\list.jpg")
    Image2(4).Picture = LoadPicture(App.path & "\images\borrow.jpg")
    Image2(5).Picture = LoadPicture(App.path & "\images\return.jpg")
    Image2(6).Picture = LoadPicture(App.path & "\images\logout.jpg")
    Image2(7).Picture = LoadPicture(App.path & "\images\exit.jpg")
    Image2(8).Picture = LoadPicture(App.path & "\images\chroma.jpg")
    
    Call SetMenuIcon
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

Private Sub opc_Click()
    Dim path As String
Dim file As String

'path = "notepad.exe"
file = App.path & "\About.pdf"
'Shell path & " " & file, vbNormalFocus

OpenChrome file
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
