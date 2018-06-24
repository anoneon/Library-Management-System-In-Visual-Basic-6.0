VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form listBook 
   BackColor       =   &H00404040&
   Caption         =   "BOOK"
   ClientHeight    =   8730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   14535
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      TabIndex        =   19
      Top             =   6480
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      TabIndex        =   18
      Top             =   5760
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      TabIndex        =   17
      Top             =   5040
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      TabIndex        =   16
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text3listBk 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command2listBk 
      BackColor       =   &H00FF8080&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8040
      Width           =   1815
   End
   Begin VB.ComboBox Combo1listBk 
      Height          =   315
      Left            =   5160
      TabIndex        =   0
      Text            =   "By"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1listBk 
      Height          =   285
      Left            =   7080
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command1listBk 
      BackColor       =   &H00808080&
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin LBMSuansu.AutoResize Resize 
      Left            =   14040
      Tag             =   "NO"
      Top             =   8040
      _ExtentX        =   714
      _ExtentY        =   714
      AspectRatioValue=   0
   End
   Begin MSDataGridLib.DataGrid DataGrid1listBk 
      Height          =   6615
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   11668
      _Version        =   393216
      AllowArrows     =   -1  'True
      BackColor       =   0
      ForeColor       =   65280
      HeadLines       =   2
      RowHeight       =   13
      RowDividerStyle =   5
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   9840
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Available Copies"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9000
      TabIndex        =   14
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9000
      TabIndex        =   13
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9000
      TabIndex        =   12
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Author"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9000
      TabIndex        =   11
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9000
      TabIndex        =   10
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Shape Shape1listBk 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   6855
      Left            =   0
      Top             =   1080
      Width           =   14535
   End
   Begin VB.Label Label4listBk 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1listBk 
      BackStyle       =   0  'Transparent
      Caption         =   "List Of Books"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label2listBk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search...."
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   0
      X2              =   3240
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   3240
      X2              =   4080
      Y1              =   600
      Y2              =   1080
   End
End
Attribute VB_Name = "listBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim ct As Integer
Sub p2()
    
    Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open "select * from Books", cn, adOpenDynamic, adLockPessimistic

    
        Set DataGrid1listBk.DataSource = rs
        DataGrid1listBk.Refresh
        
End Sub

Private Sub Combo1listBk_DropDown()
    
    Combo1listBk.Clear
    Combo1listBk.AddItem "ISBN"
    Combo1listBk.AddItem "Title"
    Combo1listBk.AddItem "Author"
    Combo1listBk.AddItem "Subject"
    
End Sub

Private Sub Command1_Click()
    
    If Text1listBk.Text = "" Then
        MsgBox "Search string empty", vbCritical, "Errorrr"
    Else
        If Combo1listBk.Text = "ISBN" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1listBk.Refresh
            rs.Open "select * from Books where ISBN like  '" & Text1listBk.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "Not Found", vbCritical, "LBM"
            Else
                Set DataGrid1listBk.DataSource = rs
            End If
        
        ElseIf Combo1listBk.Text = "Title" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1listBk.Refresh
            rs.Open "select * from Books where Title like  '" & Text1listBk.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "Not Found", vbCritical, "LBM"
            Else
                Set DataGrid1listBk.DataSource = rs
            End If
        
        ElseIf Combo1listBk.Text = "Author" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1listBk.Refresh
            rs.Open "select * from Books where Author like  '" & Text1listBk.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "Not Found", vbCritical, "LBM"
            Else
                Set DataGrid1listBk.DataSource = rs
            End If
            
        ElseIf Combo1listBk.Text = "Subject" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1listBk.Refresh
            rs.Open "select * from Books where Subject like  '" & Text1listBk.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "Not Found", vbCritical, "LBM"
            Else
                Set DataGrid1listBk.DataSource = rs
            End If
        End If
    End If

    
End Sub

Private Sub Command1listBk_Click()
    
    Call p2
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text1listBk.Text = ""
    Combo1listBk.Text = "By"
    Image1.Picture = LoadPicture("")
End Sub

Private Sub Command2listBk_Click()
    
    Unload Me
    mainWindow.Show
End Sub



Private Sub DataGrid1listBk_SelChange(Cancel As Integer)
    
     If rs.RecordCount = 0 Then
        MsgBox "Empty BOOKS", vbCritical, "LBM"
        Exit Sub
        listBook.Show
    Else
        Text1.Text = DataGrid1listBk.Columns(1).Text
        Text2.Text = DataGrid1listBk.Columns(2).Text
        Text3.Text = DataGrid1listBk.Columns(5).Text
        Text5.Text = DataGrid1listBk.Columns(7).Text
        Text4.Text = DataGrid1listBk.Columns(12).Text
        Image1.Picture = LoadPicture(DataGrid1listBk.Columns(13).Text)
    End If
End Sub

Private Sub Form_Load()
    
    ct = 0
    Me.Icon = LoadPicture(App.path & "\images\lbm_ico.ico")
    
    Set cn = New ADODB.Connection
    cn.Provider = "microsoft.jet.oledb.4.0"
    cn.Open App.path & "\dbase\dBase.mdb"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "select * from Books order by ISBN", cn, adOpenDynamic, adLockPessimistic

    
    Set DataGrid1listBk.DataSource = rs
    DataGrid1listBk.Refresh

    Text3listBk.Text = rs.RecordCount
    
    
End Sub
