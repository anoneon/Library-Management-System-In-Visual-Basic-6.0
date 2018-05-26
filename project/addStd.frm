VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form addStd 
   BackColor       =   &H00808080&
   Caption         =   "STUDENT"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton command8addStd 
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
      TabIndex        =   20
      Top             =   7920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
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
      Height          =   405
      Left            =   3720
      TabIndex        =   11
      Text            =   "Field"
      Top             =   7080
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   6120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   8421504
      Format          =   95027201
      CurrentDate     =   43190
   End
   Begin VB.TextBox Text10addStd 
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
      Left            =   10680
      TabIndex        =   15
      Top             =   6120
      Width           =   2175
   End
   Begin VB.TextBox Text11addStd 
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
      Left            =   10680
      TabIndex        =   16
      Top             =   6600
      Width           =   2175
   End
   Begin VB.ComboBox Combo2addStd 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   10680
      TabIndex        =   13
      Text            =   "Year/Batch"
      Top             =   5160
      Width           =   2415
   End
   Begin VB.ComboBox Combo1addStd 
      Height          =   315
      Left            =   5880
      TabIndex        =   23
      Text            =   "By"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Text1addStd 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      TabIndex        =   22
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Command1addStd 
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
      Left            =   8040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text2addStd 
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
      Height          =   405
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text4addStd 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   5040
      Width           =   2415
   End
   Begin VB.TextBox Text6addStd 
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
      Left            =   3720
      TabIndex        =   10
      Top             =   6600
      Width           =   2415
   End
   Begin VB.TextBox Text8addStd 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   10680
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   4440
      Width           =   3495
   End
   Begin VB.TextBox Text9addStd 
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
      Left            =   10680
      TabIndex        =   14
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox Text3addStd 
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
      Left            =   3720
      TabIndex        =   7
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton command3addStd 
      BackColor       =   &H0000FF00&
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
      TabIndex        =   6
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton command4addStd 
      BackColor       =   &H0000FFFF&
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
      TabIndex        =   5
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton command5addStd 
      BackColor       =   &H00FF0000&
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
      TabIndex        =   4
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton command6addStd 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   3
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton command9addStd 
      BackColor       =   &H000080FF&
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
      TabIndex        =   2
      Top             =   7920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton command10addStd 
      BackColor       =   &H000000FF&
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
      TabIndex        =   1
      Top             =   7920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer Timer1addStd 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   13440
      Top             =   1200
   End
   Begin VB.CommandButton Command2addStd 
      BackColor       =   &H00FF0000&
      Caption         =   "&FIND"
      Height          =   375
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton command7addStd 
      BackColor       =   &H000000FF&
      Caption         =   "Click To Upload"
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
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7080
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13200
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1addStd 
      Height          =   3255
      Left            =   360
      TabIndex        =   17
      Top             =   1200
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
   Begin Project1.AutoResize Resize 
      Left            =   12480
      Tag             =   "NO"
      Top             =   7920
      _ExtentX        =   714
      _ExtentY        =   714
      AspectRatioValue=   0
   End
   Begin VB.Label Label9addStd 
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
      Height          =   375
      Left            =   7440
      TabIndex        =   37
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Label Label14addStd 
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
      Left            =   7440
      TabIndex        =   27
      Top             =   6600
      Width           =   2775
   End
   Begin VB.Image Image1addStd 
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   10560
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label5addStd 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Name of Student"
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
      Left            =   360
      TabIndex        =   34
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label Label6addStd 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Date of Birth"
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
      Left            =   360
      TabIndex        =   33
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Label Label7addStd 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Phone No."
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
      Left            =   360
      TabIndex        =   32
      Top             =   6600
      Width           =   3015
   End
   Begin VB.Label Label8addStd 
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
      Height          =   375
      Left            =   360
      TabIndex        =   31
      Top             =   7080
      Width           =   3015
   End
   Begin VB.Label Label4addStd 
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
      Height          =   375
      Left            =   360
      TabIndex        =   26
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Label Label10addStd 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Address"
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
      Left            =   7440
      TabIndex        =   30
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label Label12addStd 
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
      Left            =   7440
      TabIndex        =   29
      Top             =   5640
      Width           =   2775
   End
   Begin VB.Label Label13addStd 
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
      Left            =   7440
      TabIndex        =   28
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Label Label15addStd 
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
      Left            =   7440
      TabIndex        =   24
      Top             =   7080
      Width           =   2775
   End
   Begin VB.Shape Shape1addBk 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   6615
      Left            =   120
      Top             =   1080
      Width           =   14295
   End
   Begin VB.Label Label2addStd 
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
      Height          =   255
      Left            =   5880
      TabIndex        =   36
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3addStd 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Total Student"
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
      Left            =   12240
      TabIndex        =   35
      Top             =   120
      Width           =   1815
   End
   Begin VB.Line Line1addStd 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   0
      X2              =   3840
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1addStd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Student Registration"
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
      Height          =   615
      Left            =   -360
      TabIndex        =   25
      Top             =   0
      Width           =   4335
   End
   Begin VB.Line Line2addStd 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   3840
      X2              =   4680
      Y1              =   600
      Y2              =   1080
   End
End
Attribute VB_Name = "addStd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim st As String
Dim tmp, tmp1, tmp2
Sub ref()
    
    Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open "select * from Std", cn, adOpenDynamic, adLockPessimistic
        
    
        Set DataGrid1addStd.DataSource = rs
        DataGrid1addStd.Refresh
        Text2addStd.Text = rs.RecordCount
End Sub

Private Sub Combo1_DropDown()
    
    If Text6addStd.Text = "" Then
        MsgBox "Enter phone No."
        Text6addStd.SetFocus
        Combo1.Text = "Field"
    Else
        Combo1.Clear
        Combo1.AddItem "BCA"
        Combo1.AddItem "Bsc Computer science"
        Combo1.AddItem "Bsc Mathematics"
        Combo1.AddItem "Bsc Chemistry"
        Combo1.AddItem "Bsc Physics"
        Combo1.AddItem "Bsc Botany"
        Combo1.AddItem "Bsc Zoology"
        Combo1.AddItem "BA Economics"
        Combo1.AddItem "BA History"
        Combo1.AddItem "BA Political Science"
        Combo1.AddItem "BA Geography"
    End If
End Sub

Private Sub Combo1addStd_DropDown()
    
    Combo1addStd.Clear
    Combo1addStd.AddItem "ID"
    Combo1addStd.AddItem "Name"
    Combo1addStd.AddItem "Year"
End Sub

Private Sub Combo2addStd_DropDown()
    
If Text8addStd.Text = "" Then
    MsgBox "Address cant be null"
    Text8addStd.SetFocus
    Text8addStd.Text = ""
    Combo2addStd.Text = "Year/Batch"
ElseIf Len(Text8addStd.Text) < 10 Then
    MsgBox "Address cant be less than 10 charecter"
    Text8addStd.SetFocus
    Text8addStd.Text = ""
    Combo2addStd.Text = "Year/Batch"
Else
    Combo2addStd.Clear
    Combo2addStd.AddItem "1st Year"
    Combo2addStd.AddItem "2nd Year"
    Combo2addStd.AddItem "3rd Year"
    Combo2addStd.AddItem "4th Year"
End If
End Sub

Private Sub command10addStd_Click()
       
        If Text6addStd.Text = tmp1 And Text8addStd.Text = tmp2 Then
            MsgBox ("No changes Done")
            Call command9addStd_Click
            Exit Sub
        ElseIf Text6addStd.Text = "" Or Text8addStd.Text = "" Then
            MsgBox ("NO Empty field")
            addStd.Show
            Exit Sub
        ElseIf Text6addStd.Text <> tmp1 Or Text8addStd.Text <> tmp2 Then
            MsgBox ("change will be done")
        End If
        
        
        tmp = MsgBox("Are You Sure", vbYesNo + vbQuestion)
        If tmp = vbYes Then
            cn.Execute ("update Std set Phone='" & Text6addStd.Text & "',Address='" & Text8addStd.Text & "' where ID='" & Text3addStd.Text & "'")
            MsgBox ("Student data Updated")
            Unload Me
            addStd.Show
            
        Else
            
            addStd.Show
        End If
        
End Sub

Private Sub Command1addStd_Click(Index As Integer)
    
    Call ref
    Combo1addStd.Text = "By"
    Text1addStd.Text = ""
    Image1addStd.Picture = LoadPicture("")
    Call command9addStd_Click
    
End Sub

Private Sub Command2addStd_Click()
    If Text1addStd.Text = "" Then
        MsgBox "Search string empty", vbCritical, "error"
    Else
        If Combo1addStd.Text = "ID" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1addStd.Refresh
            rs.Open "select * from Std where ID like  '" & Text1addStd.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "not found"
            Else
                Set DataGrid1addStd.DataSource = rs
            End If
        
        ElseIf Combo1addStd.Text = "Name" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1addStd.Refresh
            rs.Open "select * from Std where StdName like  '" & Text1addStd.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "not found"
            Else
                Set DataGrid1addStd.DataSource = rs
            End If
        
        ElseIf Combo1addStd.Text = "Year" Then
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            DataGrid1addStd.Refresh
            rs.Open "select * from Std where Year like  '" & Text1addStd.Text & "%'", cn, adOpenDynamic, adLockPessimistic
            
            If rs.EOF Then
                MsgBox "not found"
            Else
                Set DataGrid1addStd.DataSource = rs
            End If
            
        End If
    End If
End Sub

Private Sub command3addStd_Click()
    
    
    Text3addStd.Enabled = True
    Text3addStd.SetFocus
    Text4addStd.Enabled = True
    DTPicker1.Enabled = True
    Text6addStd.Enabled = True
    Combo2addStd.Enabled = True
    Combo1.Enabled = True
    Text8addStd.Enabled = True
    Text9addStd.Enabled = True
    
    Text3addStd.Text = ""
    Text4addStd.Text = ""
    Text6addStd.Text = ""
    Text8addStd.Text = ""
    Text9addStd.Text = ""
    Text10addStd.Text = ""
    Text11addStd.Text = ""
    Combo1addStd.Text = "By"
    Combo2addStd.Text = "Year/Batch"
    Combo1.Text = "Field"
    Image1addStd.Picture = LoadPicture("")
    Timer1addStd.Enabled = True
    
    command3addStd.Visible = False
    command4addStd.Enabled = False
    command5addStd.Visible = False
    command6addStd.Enabled = False
    command7addStd.Enabled = True
    command8addStd.Enabled = True
    command8addStd.Visible = True
    command9addStd.Visible = True
    command9addStd.Enabled = True
    
    command10addStd.Enabled = True
End Sub

Private Sub command4addStd_Click()
    
    Dim o
    If rs.RecordCount = 0 Then
            MsgBox "DataBase is EMPTY", vbCritical, "EMPTY"
            Exit Sub
            addStd.Show
    End If
    If Text3addStd.Text = "" Then
        MsgBox ("select a Field First")
    Else
        o = MsgBox("Are You Sure?", vbYesNo)
        If o = vbYes Then
            rs.Delete
            Call ref
            Call command9addStd_Click
        Else
            Call ref
            Call command9addStd_Click
        End If
    End If
        
End Sub

Private Sub command5addStd_Click()
    Unload Me
    mainWindow.Show
End Sub

Private Sub command6addStd_Click()
    If rs.RecordCount = 0 Then
        MsgBox ("Database Empty")
        Exit Sub
    End If
    If Text3addStd.Text = "" Then
        MsgBox ("Select a student first")
        addStd.Show
     
    ElseIf rs.RecordCount = 0 Then
        MsgBox "No Records there to Update"
        Exit Sub
        
    Else
        MsgBox ("you can only update Address and Ph no.")
        
        Combo1addStd.Text = "By"
        Combo1.Enabled = False
        command8addStd.Visible = False
        command3addStd.Enabled = False
        command5addStd.Visible = False
        command6addStd.Visible = False
        command9addStd.Visible = True
        command9addStd.Enabled = True
        command10addStd.Visible = True
        command10addStd.Enabled = True
        Text3addStd.Enabled = False
        Text4addStd.Enabled = False
        DTPicker1.Enabled = False
        Combo1.Enabled = False
        Combo2addStd.Enabled = False
        Text9addStd.Enabled = False
        Text10addStd.Enabled = False
        Text11addStd.Enabled = False
            
        Text6addStd.Enabled = True
        Text6addStd.SetFocus
        Text8addStd.Enabled = True
        
       
    End If
End Sub

Private Sub command7addStd_Click()
    
    Set rs = cn.Execute("select username from Users where username='" & Text9addStd.Text & "'")
    If Text9addStd.Text = "" Then
        MsgBox "Enter Registered Person"
        Text9addStd.SetFocus
    ElseIf rs.EOF Then
        MsgBox "Enter True Registered Person"
        Text9addStd.SetFocus
        Text9addStd.Text = ""
    Else
        CommonDialog1.ShowOpen
        CommonDialog1.Filter = "Jpeg|*.jpg"
        st = CommonDialog1.FileName
        Image1addStd.Picture = LoadPicture(st)
    End If
    
End Sub

Private Sub command8addStd_Click()
    
    Dim I As Integer
    
    If st = "" And Text9addStd.Text = "" Then
        MsgBox ("empty field")
    Else
        Set rs = cn.Execute("select * from Std where ID='" & Text3addStd.Text & "'")
    
        If rs.EOF Then
            I = MsgBox("Are you sure? ***only upadte some detail in future", vbYesNo)
            If I = vbYes Then
                cn.Execute ("insert into Std values('" & Text3addStd.Text & "','" & Text4addStd.Text & "','" & DTPicker1.Value & "','" & Text6addStd.Text & "','" & Combo1.Text & "','" & Text8addStd.Text & "','" & Combo2addStd.Text & "','" & Text9addStd.Text & "','" & Text10addStd.Text & "','" & Text11addStd.Text & "','" & st & "')")
                Call ref
                Text3addStd.Text = ""
                Text4addStd.Text = ""
                DTPicker1.Value = Date
                DTPicker1.Enabled = False
                Text6addStd.Text = ""
                Combo1.Text = "Field"
                Combo1.Enabled = False
                Text8addStd.Text = ""
                Combo2addStd.Text = "Year/Batch"
                Combo2addStd.Enabled = False
                Text9addStd.Text = ""
                Text10addStd.Text = ""
                Text11addStd.Text = ""
                Call command9addStd_Click
                st = ""
                Image1addStd.Picture = LoadPicture("")
            Else
                addStd.Show
            End If                          'VByes close
        Else                                                                'else of eof
           I = MsgBox("Student Exists", vbOKOnly)
        
            If I = vbOK Then
            
                Call ref
                Text3addStd.Text = ""
                Text4addStd.Text = ""
                DTPicker1.Value = Date
                
                Text6addStd.Text = ""
                Combo1.Text = "Field"
                
                Text8addStd.Text = ""
                Combo2addStd.Text = "Year/Batch"
                
                Text9addStd.Text = ""
                Text10addStd.Text = ""
                Text11addStd.Text = ""
                
            End If                              'vbok
        
        End If
    End If
End Sub

Private Sub command9addStd_Click()
        
    Text3addStd.Text = ""
    Text4addStd.Text = ""
    Text6addStd.Text = ""
    Text8addStd.Text = ""
    Text9addStd.Text = ""
    Text10addStd.Text = ""
    Text11addStd.Text = ""
    Combo1addStd.Text = "By"
    Combo2addStd.Text = "Year/Batch"
    Combo1.Text = "Field"
    Text3addStd.Enabled = False
    Text4addStd.Enabled = False
    DTPicker1.Enabled = False
    Text6addStd.Enabled = False
    Combo2addStd.Enabled = False
    Combo1.Enabled = False
    Text8addStd.Enabled = False
    Text9addStd.Enabled = False
    Text10addStd.Enabled = False
    Text11addStd.Enabled = False
    
    Timer1addStd.Enabled = False
    command7addStd.Enabled = False
    command3addStd.Visible = True
    command3addStd.Enabled = True
    command4addStd.Enabled = True
    command4addStd.Visible = True
    command5addStd.Enabled = True
    command5addStd.Visible = True
    command6addStd.Enabled = True
    command6addStd.Visible = True
    command8addStd.Enabled = False
    command8addStd.Visible = False
    command9addStd.Visible = False
    command9addStd.Enabled = False
    command10addStd.Visible = False
    command10addStd.Enabled = False
    
    
End Sub



Private Sub DataGrid1addStd_SelChange(Cancel As Integer)
    
    If rs.RecordCount = 0 Then
        MsgBox ("Empty database")
        Exit Sub
        addStd.Show
    End If
    Text3addStd.Text = DataGrid1addStd.Columns(0)
    Text4addStd.Text = DataGrid1addStd.Columns(1)
    DTPicker1.Value = DataGrid1addStd.Columns(2)
    Text6addStd.Text = DataGrid1addStd.Columns(3)
    Combo1.Text = DataGrid1addStd.Columns(4)
    Text8addStd.Text = DataGrid1addStd.Columns(5)
    Combo2addStd.Text = DataGrid1addStd.Columns(6)
    Text9addStd.Text = DataGrid1addStd.Columns(7)
    Text10addStd.Text = DataGrid1addStd.Columns(8)
    Text11addStd.Text = DataGrid1addStd.Columns(9)
    Image1addStd.Picture = LoadPicture(DataGrid1addStd.Columns(10))
    
    tmp1 = DataGrid1addStd.Columns(3)
    tmp2 = DataGrid1addStd.Columns(5)
    
End Sub

Private Sub DTPicker1_Change()
    If Text4addStd.Text = "" Then
        DTPicker1.Value = Date
        MsgBox "Name can't be emopty"
        Text4addStd.SetFocus
    Else
        DTPicker1.SetFocus
    End If
End Sub


Private Sub Form_Load()
    
    Me.Icon = LoadPicture(App.Path & "\images\lbm_ico.ico")
    
    
    Set cn = New ADODB.Connection
    cn.Provider = "microsoft.jet.OLEDB.4.0"
    cn.Open App.Path & "\dbase\dBase.mdb"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "select * from Std order by ID", cn, adOpenDynamic, adLockPessimistic

    
    Set DataGrid1addStd.DataSource = rs
    Text2addStd.Text = rs.RecordCount
    DataGrid1addStd.Refresh
    
    Text3addStd.Enabled = False
    Text4addStd.Enabled = False
    DTPicker1.Enabled = False
    DTPicker1.Value = Date
    Text6addStd.Enabled = False
    Combo1.Enabled = False
    Text8addStd.Enabled = False
    Combo1addStd.Text = "By"
    Combo2addStd.Text = "Year/Batch"
    Combo2addStd.Enabled = False
    Text9addStd.Enabled = False
    Text10addStd.Enabled = False
    Text11addStd.Enabled = False
End Sub

Private Sub Text3addStd_Change()
    
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
    
End Sub


Private Sub Text4addStd_GotFocus()
    If Text3addStd.Text = "" Then
        MsgBox "ID can't be Empty"
        Text3addStd.SetFocus
    Else
        Text4addStd.SetFocus
    End If
End Sub

Private Sub Text6addStd_GotFocus()
    
    If DTPicker1.Value = Date Then
        MsgBox "Set the DATE using DROPDOWN"
        DTPicker1.SetFocus
    Else
        Text6addStd.SetFocus
    End If

    
End Sub

Private Sub Text6addStd_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
    
End Sub



Private Sub Text8addStd_GotFocus()
    
    If Combo1.Text = "" Or Combo1.Text = "Field" Then
        MsgBox "Set Student Field"
        Combo1.SetFocus
    Else
        Text8addStd.SetFocus
    End If
End Sub



Private Sub Text9addStd_GotFocus()
    
    If Combo2addStd.Text = "" Or Combo2addStd.Text = "Year/Batch" Then
        MsgBox "Select Year"
        Combo2addStd.SetFocus
        Combo2addStd.Text = "Year/Batch"
    Else
        Text9addStd.SetFocus
    End If
End Sub

Private Sub Timer1addStd_Timer()
    
    Text10addStd.Text = Time
    Text11addStd.Text = Date
    
End Sub
