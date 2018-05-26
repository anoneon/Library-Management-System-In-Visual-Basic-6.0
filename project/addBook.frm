VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form addBook 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   Caption         =   "BOOK"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   13095
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Height          =   375
      Left            =   9960
      TabIndex        =   15
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Click To Upload"
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6960
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11880
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "&FIND"
      Height          =   255
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   12600
      Top             =   1080
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H0000FF00&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFF00&
      Caption         =   "Cancel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FF00FF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H000080FF&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H000080FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00FF0000&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox Text12addBk 
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
      Left            =   3360
      TabIndex        =   5
      Top             =   4560
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1addBk 
      Height          =   3255
      Left            =   120
      TabIndex        =   27
      Top             =   1080
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   5741
      _Version        =   393216
      AllowArrows     =   -1  'True
      BackColor       =   0
      Enabled         =   -1  'True
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
   Begin VB.TextBox Text10addBk 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   9960
      TabIndex        =   14
      Top             =   6000
      Width           =   2175
   End
   Begin VB.TextBox Text9addBk 
      Alignment       =   2  'Center
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
      Left            =   9960
      TabIndex        =   13
      Top             =   5520
      Width           =   2175
   End
   Begin VB.TextBox Text8addBk 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9960
      TabIndex        =   12
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox Text7addBk 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9960
      TabIndex        =   11
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Text6addBk 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   9
      Top             =   6480
      Width           =   2415
   End
   Begin VB.TextBox Text5addBk 
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
      Left            =   3360
      TabIndex        =   8
      Top             =   6000
      Width           =   2415
   End
   Begin VB.TextBox Text4addBk 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   7
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox Text3addBk 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   6
      Top             =   5040
      Width           =   2415
   End
   Begin VB.ComboBox Combo2addBk 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "addBook.frx":0000
      Left            =   3360
      List            =   "addBook.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   10
      Text            =   "Select Subject"
      Top             =   6960
      Width           =   2415
   End
   Begin Project1.AutoResize Resize 
      Left            =   12480
      Tag             =   "NO"
      Top             =   7680
      _ExtentX        =   714
      _ExtentY        =   714
      AspectRatioValue=   0
   End
   Begin VB.TextBox Text2addBk 
      Alignment       =   2  'Center
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
      Height          =   285
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1addBk 
      BackColor       =   &H00808080&
      Caption         =   "Refresh"
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
      Index           =   0
      Left            =   7680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1addBk 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7440
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.ComboBox Combo1addBk 
      Height          =   315
      Left            =   5400
      TabIndex        =   1
      Text            =   "By"
      Top             =   600
      Width           =   1575
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   9480
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Price(in INR)"
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
      Height          =   375
      Left            =   6960
      TabIndex        =   39
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Upload Pic"
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
      Height          =   375
      Left            =   6960
      TabIndex        =   38
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Book Registration"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   37
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label14addBk 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "ISBN"
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
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Line Line1addBk 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      X1              =   0
      X2              =   3600
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label13addBk 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Registered Date"
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
      Height          =   375
      Left            =   6960
      TabIndex        =   35
      Top             =   6000
      Width           =   2775
   End
   Begin VB.Label Label12addBk 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Registered Time"
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
      Height          =   375
      Left            =   6960
      TabIndex        =   34
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Label11addBk 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Registered By"
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
      Height          =   375
      Left            =   6960
      TabIndex        =   33
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label Label9addBk 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Total Copies"
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
      Height          =   375
      Left            =   6960
      TabIndex        =   32
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label Label8addBk 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Subject"
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
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   6960
      Width           =   3015
   End
   Begin VB.Label Label7addBk 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Publisher"
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
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label Label6addBk 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Copyright ©"
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
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   6000
      Width           =   3015
   End
   Begin VB.Label Label5addBk 
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
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   5520
      Width           =   3015
   End
   Begin VB.Label Label4addBk 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Title of Book"
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
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label Label3addBk 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Total"
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
      Left            =   11160
      TabIndex        =   24
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2addBk 
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
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Shape Shape1addBk 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   6615
      Left            =   0
      Top             =   960
      Width           =   13095
   End
   Begin VB.Line Line2addBk 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      X1              =   3600
      X2              =   4320
      Y1              =   720
      Y2              =   1080
   End
End
Attribute VB_Name = "addBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim st As String
Dim avl As Integer
Dim Y As String
Dim tmp As Integer
Dim l, m, a, c, s
Sub p1()
    
    Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open "select * from Books order by ISBN", cn, adOpenDynamic, adLockPessimistic

    
        Set DataGrid1addBk.DataSource = rs
        DataGrid1addBk.Refresh
        Text2addBk.Text = rs.RecordCount
        
End Sub

Private Sub cmdAdd_Click()

    Text12addBk.Enabled = True
    Text12addBk.SetFocus
    Text3addBk.Enabled = True
    Text4addBk.Enabled = True
    Text5addBk.Enabled = True
    Text6addBk.Enabled = True
    Combo2addBk.Enabled = True
    Text7addBk.Enabled = True
    Text8addBk.Enabled = True
    Text1addBk.Enabled = True
    
    Text12addBk.Text = ""
    Text1.Text = ""
    Text3addBk.Text = ""
    Text4addBk.Text = ""
    Text5addBk.Text = ""
    Text6addBk.Text = ""
    Text7addBk.Text = ""
    Text8addBk.Text = ""
    Text9addBk.Text = ""
    Text10addBk.Text = ""
    Image1.Picture = LoadPicture("")
    
    
    Combo1addBk.Text = "By"
    Combo2addBk.Text = "Select Subject"
    Text9addBk.Enabled = False
    Text10addBk.Enabled = False
    Text1.Enabled = True
    Timer1.Enabled = True
    Command2.Enabled = True
    cmdAdd.Visible = False
    cmdSave.Visible = True
    cmdDel.Enabled = False
    cmdClose.Visible = False
    cmdCancel.Visible = True
    cmdCancel.Enabled = True
    cmdEdit.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    
    Timer1.Enabled = False
    
    Text12addBk.Text = ""
    Text12addBk.Enabled = False
    Text1.Text = ""
    Text1.Enabled = False
    Text3addBk.Text = ""
    Text4addBk.Text = ""
    Text5addBk.Text = ""
    Text6addBk.Text = ""
    Text7addBk.Text = ""
    Text8addBk.Text = ""
    Combo1addBk.Text = "By"
    Combo2addBk.Text = "Select Subject"
    Combo2addBk.Enabled = False
    Text9addBk.Text = ""
    Text10addBk.Text = ""
    
    Image1.Picture = LoadPicture("")
    Text3addBk.Enabled = False
    Text4addBk.Enabled = False
    Text5addBk.Enabled = False
    Text6addBk.Enabled = False
    Text7addBk.Enabled = False
    Text8addBk.Enabled = False
   
    Text9addBk.Enabled = False
    Text10addBk.Enabled = False
    Command2.Enabled = False
    
    cmdCancel.Visible = False
    cmdClose.Visible = True
    cmdAdd.Visible = True
    cmdAdd.Enabled = True
    cmdSave.Visible = False
    cmdDel.Enabled = True
    cmdEdit.Enabled = True
    cmdEdit.Visible = True
    cmdUpdate.Visible = False
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    mainWindow.Show
End Sub

Private Sub cmdDel_Click()
        Dim k As Integer
        If rs.RecordCount = 0 Then
            MsgBox "DataBase is EMPTY", vbCritical, "EMPTY"
            Exit Sub
        End If
        If Text12addBk.Text = "" Then
            MsgBox "select the book from table"
        Else
            k = MsgBox("are you sure?", vbYesNo)
            If k = vbYes Then
                rs.Delete
                Call p1
                Call cmdCancel_Click
            Else
                Call p1
                Call cmdCancel_Click
            End If
        End If
End Sub

Sub cmdEdit_Click()
    
    If Text12addBk.Text = "" Then
        MsgBox "Select A Field"
        addBook.Show
    ElseIf rs.RecordCount = 0 Then
        MsgBox "No Records there to Update"
        Exit Sub
    Else
        MsgBox ("you can only edit total copies of book")
        
        
        
        Combo1addBk.Text = "By"
        Combo2addBk.Enabled = False
        Command2.Enabled = False
        Timer1.Enabled = True
        cmdCancel.Visible = True
        cmdCancel.Enabled = True
        cmdClose.Visible = False
        cmdAdd.Enabled = False
        cmdSave.Visible = False
        cmdDel.Enabled = True
        cmdUpdate.Visible = True
        Text7addBk.Enabled = True
        cmdEdit.Visible = False
                
        Text7addBk.Enabled = True
        Text7addBk.SetFocus
    End If
    
End Sub

Private Sub cmdSave_Click()
    Dim I, J As Integer
    Y = "YES"
    
    Set rs = cn.Execute("select * from Books where ISBN='" & Text12addBk.Text & "'")
    If st = "" Then
        MsgBox "Enter the data in given Fields"
    
    ElseIf rs.EOF Then
      J = MsgBox("are you sure? *can only update total copies in future", vbYesNo)
        
        If J = vbYes Then
            cn.Execute ("insert into Books values('" & Text12addBk.Text & "','" & Text3addBk.Text & "','" & Text4addBk.Text & "','" & Text5addBk.Text & "','" & Text6addBk.Text & "','" & Combo2addBk.Text & "','" & Text7addBk.Text & "','" & Text7addBk.Text & "','" & Y & "','" & Text8addBk.Text & "','" & Text10addBk.Text & "','" & Text9addBk.Text & "','" & Text1.Text & "','" & st & "')")
        
            I = MsgBox("Book Added", 0 + vbInformation, "ADDED")
        
            If I = vbOK Then
                Call p1
                Text12addBk.Text = ""
                Text3addBk.Text = ""
                Text4addBk.Text = ""
                Text5addBk.Text = ""
                Text6addBk.Text = ""
                Text7addBk.Text = ""
                Text8addBk.Text = ""
                Combo1addBk.Text = "By"
                Combo2addBk.Text = "Select Subject"
                Combo2addBk.Enabled = False
                Text9addBk.Text = ""
                Text10addBk.Text = ""
                Text1.Text = ""
                Call cmdCancel_Click
            
            End If
        Else
            addBook.Show
        End If
        
    Else
        I = MsgBox("Book Already Exists", 0 + vbCritical)
        If I = vbOK Then
            Call p1
            Text12addBk.Text = ""
            Text12addBk.SetFocus
            Text3addBk.Text = ""
            Text4addBk.Text = ""
            Text5addBk.Text = ""
            Text6addBk.Text = ""
            Text7addBk.Text = ""
            Text8addBk.Text = ""
            Combo2addBk.Text = "Select Subject"
            Text9addBk.Text = ""
            Text10addBk.Text = ""
            Text1.Text = ""
        End If
    End If
    
    
End Sub

Private Sub cmdUpdate_Click()
        
        If rs.RecordCount = 0 Then
        MsgBox ("Database Empty")
        Exit Sub
        End If
        If Text7addBk.Text <> l Then
            m = Text7addBk.Text
            a = m - l
            c = DataGrid1addBk.Columns(7) + a
        ElseIf Text7addBk.Text = l Then
            MsgBox ("None Books are added")
            Call cmdCancel_Click
            Exit Sub
        End If
        
        If a > 0 Then
         MsgBox (a & " books will be added")
        Else
         MsgBox (a & "books will be decreased")
        End If
        
        
        tmp = MsgBox("Are You Sure", vbYesNo + vbQuestion)
        If tmp = vbYes Then
            cn.Execute ("update Books set Total_Copies='" & Text7addBk & "',Available_Copies='" & c & "' where ISBN='" & Text12addBk.Text & "'")
            MsgBox ("Book Updated")
            Unload Me
            addBook.Show
            
        Else
            
            addBook.Show
        End If
   
End Sub

Private Sub Combo1addBk_DropDown()
    
        Combo1addBk.Clear
        Combo1addBk.AddItem "ISBN"
        Combo1addBk.AddItem "Title"
        Combo1addBk.AddItem "Author"
        Combo1addBk.AddItem "Subject"
    
    
End Sub


Private Sub Combo2addBk_DropDown()
    
     If Text6addBk.Text = "" Or Text6addBk.Text = "Select Subject" Then
        MsgBox "Publisher can't be emopty"
        Combo2addBk.Text = "Select Subject"
        Text6addBk.SetFocus
    Else
        Combo2addBk.Clear
        Combo2addBk.AddItem "Novel"
        Combo2addBk.AddItem "Mathematics"
        Combo2addBk.AddItem "Chemistry"
        Combo2addBk.AddItem "Physics"
        Combo2addBk.AddItem "Biology"
        Combo2addBk.AddItem "Computer"
    End If
    
End Sub

Private Sub Combo2addBk_GotFocus()
    
    Call Combo2addBk_DropDown
    
    
    
    
End Sub

Private Sub Command1_Click()
    
    If Text1addBk.Text = "" Then
        MsgBox "Search string empty", vbCritical, "error"
    Else
        If Combo1addBk.Text = "ISBN" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1addBk.Refresh
            rs.Open "select * from Books where ISBN like  '" & Text1addBk.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "not found"
            Else
                Set DataGrid1addBk.DataSource = rs
            End If
        
        ElseIf Combo1addBk.Text = "Title" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1addBk.Refresh
            rs.Open "select * from Books where Title like  '" & Text1addBk.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "not found"
            Else
                Set DataGrid1addBk.DataSource = rs
            End If
        
        ElseIf Combo1addBk.Text = "Author" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1addBk.Refresh
            rs.Open "select * from Books where Author like  '" & Text1addBk.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "not found"
            Else
                Set DataGrid1addBk.DataSource = rs
            End If
            
        ElseIf Combo1addBk.Text = "Subject" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1addBk.Refresh
            rs.Open "select * from Books where Subject like  '" & Text1addBk.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "not found"
            Else
                Set DataGrid1addBk.DataSource = rs
            End If
        End If
    End If
End Sub

Private Sub Command1addBk_Click(Index As Integer)
  
    Call p1
    Text12addBk.Text = ""
    Text3addBk.Text = ""
    Text4addBk.Text = ""
    Text5addBk.Text = ""
    Text6addBk.Text = ""
    Text7addBk.Text = ""
    Text8addBk.Text = ""
    Combo1addBk.Text = "By"
    Combo2addBk.Text = "Select Subject"
    Text9addBk.Text = ""
    Text10addBk.Text = ""
    Text1.Text = ""
    Text1addBk.Text = ""
    Image1.Picture = LoadPicture("")
    
End Sub

Private Sub Command2_Click()
    
    
    
    If Text1.Text = "" Then
        MsgBox "Enter Price"
        Text1.SetFocus
    Else
        CommonDialog1.FileName = ""
        CommonDialog1.Filter = "Jpeg|*.jpg"
        CommonDialog1.ShowOpen
        
'**** conditoion for samll size
        st = CommonDialog1.FileName
        Image1.Picture = LoadPicture(st)
    End If
End Sub

Private Sub DataGrid1addBk_SelChange(Cancel As Integer)

    If Not rs.EOF Then
        Text12addBk.Text = DataGrid1addBk.Columns(0).Text
        Text3addBk.Text = DataGrid1addBk.Columns(1).Text
        Text4addBk.Text = DataGrid1addBk.Columns(2).Text
        Text5addBk.Text = DataGrid1addBk.Columns(3).Text
        Text6addBk.Text = DataGrid1addBk.Columns(4).Text
        Combo2addBk.Text = DataGrid1addBk.Columns(5).Text
        Text7addBk.Text = DataGrid1addBk.Columns(6).Text
        Text8addBk.Text = DataGrid1addBk.Columns(9).Text
        Text9addBk.Text = DataGrid1addBk.Columns(10).Text
        Text10addBk.Text = DataGrid1addBk.Columns(11).Text
        Text1.Text = DataGrid1addBk.Columns(12).Text
        Image1.Picture = LoadPicture(DataGrid1addBk.Columns(13).Text)
        
     End If
     l = Text7addBk.Text
End Sub

Private Sub Form_Load()
    
    Me.Icon = LoadPicture(App.Path & "\images\lbm_ico.ico")
    
    Command2.Enabled = False
    
    
    Set cn = New ADODB.Connection
    cn.Provider = "microsoft.jet.OLEDB.4.0"
    cn.Open App.Path & "\dbase\dBase.mdb"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "select * from Books order by ISBN", cn, adOpenDynamic, adLockPessimistic

    
    Set DataGrid1addBk.DataSource = rs
    DataGrid1addBk.Refresh
    
    Text2addBk.Text = rs.RecordCount
    
    
    Text12addBk.Enabled = False
    Text3addBk.Enabled = False
    Text4addBk.Enabled = False
    Text5addBk.Enabled = False
    Text6addBk.Enabled = False
    Text7addBk.Enabled = False
    Text8addBk.Enabled = False
    Combo1addBk.Text = "By"
    Combo2addBk.Text = "Select Subject"
    Combo2addBk.Enabled = False
    Text9addBk.Enabled = False
    Text10addBk.Enabled = False
    

   
   
'  available cannot be greater than total


End Sub

Private Sub Text1_GotFocus()
    
Set rs = cn.Execute("select username from Users where username='" & Text8addBk.Text & "'")
    If Text8addBk.Text = "" Then
        MsgBox "Registered By can't be emopty"
        Text8addBk.SetFocus
    ElseIf rs.EOF Then
        MsgBox "Registered user doesn't match"
        Text8addBk.SetFocus
        Text8addBk.Text = ""
    Else
        Text1.SetFocus
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
    
End Sub

Private Sub Text3addBk_GotFocus()
    
        If Text12addBk.Text = "" Then
            MsgBox "ISBN can't be emopty"
            Text12addBk.SetFocus
        Else
            Text3addBk.SetFocus
        End If
    
    End Sub



Private Sub Text4addBk_GotFocus()
    
    If Text3addBk.Text = "" Then
        MsgBox "Title can't be emopty"
        Text3addBk.SetFocus
    Else
        Text4addBk.SetFocus
    End If
    
End Sub
Private Sub Text5addBk_GotFocus()
    
    If Text4addBk.Text = "" Then
        MsgBox "Author can't be emopty"
        Text4addBk.SetFocus
    Else
        Text5addBk.SetFocus
    End If
    
End Sub
Private Sub Text6addBk_GotFocus()
    
    If Text5addBk.Text = "" Then
        MsgBox "Copyright can't be emopty"
        Text5addBk.SetFocus
    Else
        Text6addBk.SetFocus
    End If
    
End Sub



Private Sub Text7addBk_GotFocus()
       If Combo2addBk.Text = "Select Subject" Or Combo2addBk.Text = "" Then
        MsgBox "Select Some Subject"
        Combo2addBk.SetFocus
        Combo2addBk.Text = "Select Subject"
    ElseIf Not (Combo2addBk.Text = "Select Subject" Or Combo2addBk.Text = "Novel" Or Combo2addBk.Text = "Mathematics" Or Combo2addBk.Text = "Chemistry" Or Combo2addBk.Text = "Physics" Or Combo2addBk.Text = "Biology" Or Combo2addBk.Text = "Computer") Then
        MsgBox "cant be that"
        Combo2addBk.SetFocus
    Else
        Text7addBk.SetFocus
    End If
End Sub

Private Sub Text7addBk_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub


Private Sub Text8addBk_GotFocus()
    
    If Text7addBk.Text = "" Then
        MsgBox "Total Copies can't be emopty"
        Text7addBk.SetFocus
    Else
        Text8addBk.SetFocus
    End If
    
End Sub

Private Sub Timer1_Timer()
    
    Text9addBk.Text = Time
    Text10addBk.Text = Date
    
    
End Sub
