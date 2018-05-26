VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form listStd 
   BackColor       =   &H00404040&
   Caption         =   "STUDENT"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   14505
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1listStd 
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
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1listStd 
      Height          =   285
      Left            =   7080
      TabIndex        =   8
      Top             =   600
      Width           =   2055
   End
   Begin VB.ComboBox Combo1listStd 
      Height          =   315
      Left            =   5160
      TabIndex        =   7
      Text            =   "By"
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command2listStd 
      BackColor       =   &H00FF0000&
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
      TabIndex        =   6
      Top             =   8040
      Width           =   1815
   End
   Begin VB.TextBox Text3listStd 
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
      Left            =   13080
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   495
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
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text1 
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
      Height          =   495
      Left            =   12000
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox Text2 
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
      Height          =   975
      Left            =   12000
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox Text3 
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
      Height          =   495
      Left            =   12000
      TabIndex        =   1
      Top             =   5640
      Width           =   2415
   End
   Begin VB.TextBox Text4 
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
      Height          =   495
      Left            =   12000
      TabIndex        =   0
      Top             =   6360
      Width           =   2415
   End
   Begin Project1.AutoResize Resize 
      Left            =   14040
      Tag             =   "NO"
      Top             =   8040
      _ExtentX        =   714
      _ExtentY        =   714
      AspectRatioValue=   0
   End
   Begin MSDataGridLib.DataGrid DataGrid1listStd 
      Height          =   6615
      Left            =   120
      TabIndex        =   10
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "ID"
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
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Name"
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
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Field"
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
      Top             =   5640
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Year"
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
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   10560
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   0
      X2              =   3600
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2listStd 
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
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1listStd 
      BackStyle       =   0  'Transparent
      Caption         =   "List Of Students"
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
      TabIndex        =   16
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label4listStd 
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
      Left            =   12480
      TabIndex        =   15
      Top             =   120
      Width           =   1695
   End
   Begin VB.Shape Shape1listStd 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   6855
      Left            =   0
      Top             =   1080
      Width           =   14535
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   3600
      X2              =   4080
      Y1              =   600
      Y2              =   1080
   End
End
Attribute VB_Name = "listStd"
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
        rs.Open "select * from Std", cn, adOpenDynamic, adLockPessimistic

    
        Set DataGrid1listStd.DataSource = rs
        DataGrid1listStd.Refresh
        
End Sub

Private Sub Combo1listStd_DropDown()
    
    Combo1listStd.Clear
    Combo1listStd.AddItem "ID"
    Combo1listStd.AddItem "Name"
    Combo1listStd.AddItem "Year"
    Combo1listStd.AddItem "Field"
    
End Sub

Private Sub Command1_Click()
    
    If Text1listStd.Text = "" Then
        MsgBox "Search string empty", vbCritical, "error"
    Else
        If Combo1listStd.Text = "ID" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1listStd.Refresh
            rs.Open "select * from Std where ID like  '" & Text1listStd.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "not found"
            Else
                Set DataGrid1listStd.DataSource = rs
            End If
        
        ElseIf Combo1listStd.Text = "Name" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1listStd.Refresh
            rs.Open "select * from Std where StdName like  '" & Text1listStd.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "not found"
            Else
                Set DataGrid1listStd.DataSource = rs
            End If
        
        ElseIf Combo1listStd.Text = "Year" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1listStd.Refresh
            rs.Open "select * from Std where Year like  '" & Text1listStd.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "not found"
            Else
                Set DataGrid1listStd.DataSource = rs
            End If
            
        ElseIf Combo1listStd.Text = "Field" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1listStd.Refresh
            rs.Open "select * from Std where Field like  '" & Text1listStd.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "not found"
            Else
                Set DataGrid1listStd.DataSource = rs
            End If
        End If
    End If

    
End Sub

Private Sub Command1listStd_Click()
    
    Call p2
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text1listStd.Text = ""
    Combo1listStd.Text = "By"
    Image1.Picture = LoadPicture("")
End Sub

Private Sub Command2listStd_Click()
    
    Unload Me
    mainWindow.Show
    
End Sub



Private Sub DataGrid1listStd_SelChange(Cancel As Integer)
     If rs.RecordCount = 0 Then
        MsgBox ("empty Database")
        Exit Sub
        listStd.Show
    Else
    Text1.Text = DataGrid1listStd.Columns(0).Text
    Text2.Text = DataGrid1listStd.Columns(1).Text
    Text3.Text = DataGrid1listStd.Columns(4).Text
    Text4.Text = DataGrid1listStd.Columns(6).Text
    Image1.Picture = LoadPicture(DataGrid1listStd.Columns(10).Text)
    End If
End Sub

Private Sub Form_Load()
    
    ct = 0
    Me.Icon = LoadPicture(App.Path & "\images\lbm_ico.ico")
    
    Set cn = New ADODB.Connection
    cn.Provider = "microsoft.jet.oledb.4.0"
    cn.Open App.Path & "\dbase\dBase.mdb"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "select * from Std order by ID", cn, adOpenDynamic, adLockPessimistic

    
    Set DataGrid1listStd.DataSource = rs
    DataGrid1listStd.Refresh

    Text3listStd.Text = rs.RecordCount
    
   

End Sub

