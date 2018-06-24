VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form returnBk 
   BackColor       =   &H00404040&
   Caption         =   "BORROW"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   13170
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   10560
      Top             =   7680
   End
   Begin LBMSuansu.AutoResize Resize 
      Left            =   12000
      Tag             =   "NO"
      Top             =   7800
      _ExtentX        =   714
      _ExtentY        =   714
      AspectRatioValue=   0
   End
   Begin VB.CommandButton Command4rbrw 
      BackColor       =   &H00808080&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton Command3rbrw 
      BackColor       =   &H0000FF00&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command1rbrw 
      BackColor       =   &H000040C0&
      Caption         =   "GO"
      Height          =   375
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text5rbrw 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11520
      TabIndex        =   11
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text4brw 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9480
      TabIndex        =   10
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text3rbrw 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   960
      Width           =   1935
   End
   Begin VB.ComboBox Combo1rbrw 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6480
      TabIndex        =   8
      Text            =   "By"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1rbrw 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4800
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.Frame Frame1rbrw 
      Caption         =   "RETURN SECTION : - "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   9735
      Begin VB.TextBox Text9brw 
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
         Height          =   735
         Left            =   6480
         TabIndex        =   23
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF00&
         Caption         =   "RETURN"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text8brw 
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
         Height          =   615
         Left            =   2520
         TabIndex        =   18
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox Text7brw 
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
         Height          =   615
         Left            =   2520
         TabIndex        =   16
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox Text6brw 
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
         Height          =   855
         Left            =   2520
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label10rbrw 
         Alignment       =   2  'Center
         Caption         =   "Fine"
         Height          =   495
         Left            =   5760
         TabIndex        =   22
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label9rbrw 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Return Date"
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
         Height          =   615
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label8rbrw 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Issued By"
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
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label7rbrw 
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
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   2055
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1rbrw 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   4683
      _Version        =   393216
      BackColor       =   0
      ForeColor       =   65280
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
   Begin VB.Label Label6rbrw 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   11880
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5rbrw 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9720
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4rbrw 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2rbrw 
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
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   3360
      X2              =   3960
      Y1              =   840
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   3360
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1rbrw 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Return  Book"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   -240
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   7095
      Left            =   0
      Top             =   1440
      Width           =   13215
   End
End
Attribute VB_Name = "returnBk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim tmp
Sub refres()
    Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open "select * from Borrow order by ISBN", cn, adOpenDynamic, adLockPessimistic
        
        Set DataGrid1rbrw.DataSource = rs
        DataGrid1rbrw.Refresh
        Text1rbrw.Text = rs.RecordCount
End Sub
Private Sub Combo1rbrw_DropDown()
        
    Combo1rbrw.Clear
    Combo1rbrw.AddItem "Student Name"
    Combo1rbrw.AddItem "Book ISBN"
    
End Sub

Private Sub Command1rbrw_Click()
    If Text3rbrw.Text = "" Then
        
        MsgBox "Search string empty", vbCritical, "Errorrr"
    
    Else
        
        If Combo1rbrw.Text = "Student Name" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1rbrw.Refresh
            rs.Open "select * from Borrow where StdName like  '" & Text3rbrw.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "Not found", vbExclamation, "LBM"
            Else
                Set DataGrid1rbrw.DataSource = rs
            End If
        ElseIf Combo1rbrw.Text = "Book ISBN" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1rbrw.Refresh
            rs.Open "select * from Borrow where ISBN like  '" & Text3rbrw.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "Not found", vbExclamation, "LBM"
            Else
                Set DataGrid1rbrw.DataSource = rs
            End If
            
        End If
        
    End If
    
End Sub

Private Sub Command2_Click()
Dim dummy
If rs.RecordCount = 0 Then
    MsgBox "Nothing To Return", vbInformation, "LBM"
    returnBk.Show
    Exit Sub
Else
dummy = DataGrid1rbrw.Columns(0)
End If
If Text6brw.Text = "" Then
    MsgBox "Select a data", vbExclamation, "SELECT"
    returnBk.Show
Else
        rs.Delete

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open "select * from Books", cn, adOpenDynamic, adLockPessimistic

        cn.Execute ("update Books set Available_Copies=Available_Copies+1 where ISBN='" & dummy & "'")
        MsgBox "Book returned Successfully", vbInformation, "LBM"
        Call refres
        Call Command3rbrw_Click
End If
End Sub

Private Sub Command3rbrw_Click()
    Call refres
    Text6brw.Text = ""
    Text7brw.Text = ""
    Text8brw.Text = ""
    Text9brw.Text = ""
End Sub

Private Sub Command4rbrw_Click()
    Unload Me
    mainWindow.Show
End Sub

Private Sub DataGrid1rbrw_SelChange(Cancel As Integer)
  If rs.RecordCount = 0 Then
    MsgBox "Empty Database", vbCritical, "DATABASE EMPTY"
    Exit Sub
    returnBk.Show
  Else
    Text6brw.Text = DataGrid1rbrw.Columns(1)
    Text7brw.Text = DataGrid1rbrw.Columns(6)
    Text8brw.Text = DataGrid1rbrw.Columns(5)
    
    Dim l, m, g
    
        If Date > Text8brw.Text Then
            l = Date
            m = CDate(Text8brw.Text)
            g = l - m
        If g > 0 Then
            MsgBox "fine for " & g & " days"
            Text9brw.Text = "Rupees " & g
        Else
            Text9brw.Text = "Rupees 0"
        End If
    End If
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & "\images\lbm_ico.ico")
    
    Set cn = New ADODB.Connection
    cn.Provider = "microsoft.jet.OLEDB.4.0"
    cn.Open App.path & "\dbase\dBase.mdb"
    
    Set rs = New ADODB.Recordset
    
    
    rs.CursorLocation = adUseClient
    rs.Open "select * from Borrow order by ISBN", cn, adOpenDynamic, adLockPessimistic
    
    Set DataGrid1rbrw.DataSource = rs
    
    DataGrid1rbrw.Refresh
    Text1rbrw.Text = rs.RecordCount
    
End Sub
Private Sub Timer1_Timer()
    Text4brw = Date
    Text5rbrw.Text = Time
End Sub
