VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Movable Animation?"
      Height          =   405
      Left            =   795
      TabIndex        =   6
      Top             =   6135
      Width           =   2460
   End
   Begin VB.CheckBox chkSolidBkg 
      Caption         =   "Goose Solid Bkg"
      Height          =   330
      Left            =   3705
      TabIndex        =   10
      Top             =   6240
      Width           =   2100
   End
   Begin VB.CheckBox chkModClrTbl2 
      Caption         =   "Grayscale Cheetah"
      Height          =   330
      Left            =   3705
      TabIndex        =   9
      Top             =   5850
      Width           =   2100
   End
   Begin VB.CheckBox chkModPal 
      Caption         =   "Inverse Beetle's Palette"
      Height          =   330
      Left            =   3705
      TabIndex        =   8
      Top             =   5445
      Width           =   2100
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load GIF from URL (Example)"
      Height          =   420
      Left            =   255
      TabIndex        =   5
      Top             =   5625
      Width           =   3330
   End
   Begin VB.CheckBox chkMirror 
      Caption         =   "Mirror Animation"
      Height          =   330
      Left            =   3705
      TabIndex        =   7
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   420
      Index           =   3
      Left            =   2805
      TabIndex        =   4
      ToolTipText     =   "Resume animation"
      Top             =   5085
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fwd"
      Height          =   420
      Index           =   2
      Left            =   1950
      TabIndex        =   3
      ToolTipText     =   "Moves animation to next frame"
      Top             =   5085
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pause"
      Height          =   420
      Index           =   1
      Left            =   1095
      TabIndex        =   2
      ToolTipText     =   "Pauses animation at current frame"
      Top             =   5085
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop"
      Height          =   420
      Index           =   0
      Left            =   255
      TabIndex        =   1
      ToolTipText     =   "Stops animation and displays 1st frame"
      Top             =   5085
      Width           =   780
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1800
      Left            =   210
      TabIndex        =   11
      Top             =   210
      Width           =   3255
      Begin VB.Label lblFrame 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   765
         TabIndex        =   12
         Top             =   1425
         Width           =   1800
      End
      Begin prjGIFViewer.ucAniGIF ucAniGIF 
         Height          =   900
         Index           =   0
         Left            =   525
         Top             =   480
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   1588
         GIF             =   "Form1.frx":0000
         Delay           =   45
         Loops           =   1000
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H008080FF&
      Height          =   2730
      Left            =   210
      Picture         =   "Form1.frx":1C62
      ScaleHeight     =   2670
      ScaleWidth      =   5550
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2250
      Width           =   5610
      Begin VB.Label lblFrame 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   285
         Index           =   3
         Left            =   135
         TabIndex        =   15
         Top             =   2340
         Width           =   1800
      End
      Begin VB.Label lblFrame 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   285
         Index           =   2
         Left            =   2880
         TabIndex        =   14
         Top             =   2385
         Width           =   1800
      End
      Begin prjGIFViewer.ucAniGIF ucAniGIF 
         Height          =   1260
         Index           =   3
         Left            =   240
         ToolTipText     =   "Image loaded from Remote path"
         Top             =   720
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   2223
         GIF             =   "Form1.frx":11D96
         Stretch         =   3
         Enabled         =   0   'False
      End
      Begin prjGIFViewer.ucAniGIF ucAniGIF 
         Height          =   2250
         Index           =   2
         Left            =   1980
         Top             =   105
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3969
         GIF             =   "Form1.frx":11DAE
         Loops           =   25
      End
   End
   Begin VB.Label lblFrame 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   13
      Top             =   1995
      Width           =   1800
   End
   Begin prjGIFViewer.ucAniGIF ucAniGIF 
      Height          =   1905
      Index           =   1
      Left            =   3585
      Top             =   180
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   3360
      GIF             =   "Form1.frx":36A9C
      Stretch         =   4
      Loops           =   1000
      Mirror          =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Keep these things in mind when you play.

' 1. VB is single threaded. The more animated GIFs on the form, the longer it
'   will take for all of them to begin animation.
'   - The DelayAnimation property helps the form load faster
'   - When DelayAnimation=False, form will take more time to initially display

' 2. Compile the usercontrol. If uncompiled, you can expect these common annoyances:
'   - When MsgBox pops up, images disappear until Msgbox is closed
'   - Dragging other windows over a static GIF (stopped animating) may not repaint immediately

' 3. The dotted/thick border in design time does not disappear. It is purposely painted to:
'   - Identify the animated GIF control from a normal VB image control
'   - Show you the bounds of the overall image

' 4. Do not overlap image controls if possible.
'   - Overlapped controls forward paint events to every control above it in the zOrder
'   - Having several controls overlapped theoretically can bog down your application

' 5. For performance reasons, you should pause or stop animation when your app is minimized

Option Explicit

Private Sub chkModClrTbl2_Click()

    ' Example of modifying the GIF's palette
    ' Can be useful if your GIF needs to change colors depending on user selections/actions
    
    ' This example shows how you can ask the class to keep the original palette
    ' and then ask it to restore the original palette.  Whereas the example in
    ' chkModPal event doesn't request caching of the palette, because it simply
    ' toggles color inversion.
    
    Dim P As Long, pEntries() As Long
    Dim R As Integer, G As Integer, B As Integer, E As Long
    
    If chkModClrTbl2 Then
    
        ucAniGIF(0).CacheColorTables    ' cache the original palette(s) so we can replace them later
        
        For P = 0 To ucAniGIF(0).FrameCount
            If ucAniGIF(0).GetPalette(P, pEntries()) = True Then
            
                For E = 0 To UBound(pEntries)
                    ' use simple averaging to create grayscale
                    B = (pEntries(E) And &HFF)
                    G = ((pEntries(E) \ &H100) And &HFF)
                    R = ((pEntries(E) \ &H10000) And &HFF)
                    B = (B + G + R) \ 3
                    pEntries(E) = RGB(B, B, B)
                    
                Next
                ucAniGIF(0).SetPalette P, pEntries()
            End If
        Next
    Else
    
        ucAniGIF(0).RestoreColorTables True ' restore cached palette(s) and then release memory
        If ucAniGIF(0).Action <> gfaPlay Then ucAniGIF(0).Refresh
        
    End If
    
End Sub

Private Sub chkModPal_Click()

    ' Example of modifying the GIF's palette
    ' Can be useful if your GIF needs to change colors depending on user selections/actions
    
    ' This example does not need the usercontrol to cache the original palette because
    ' when inversion is restored or activated, it is a simple XOR function on color values
    
    Dim P As Long, pEntries() As Long
    Dim R As Integer, G As Integer, B As Integer, E As Long
    
    For P = 0 To ucAniGIF(2).FrameCount
        If ucAniGIF(2).GetPalette(P, pEntries()) = True Then
        
            For E = 0 To UBound(pEntries)
                ' inverse the RGB colors. Remember that palette entires are BGR, not RGB
                B = (pEntries(E) And &HFF) Xor 255
                G = ((pEntries(E) \ &H100) And &HFF) Xor 255
                R = ((pEntries(E) \ &H10000) And &HFF) Xor 255
                
                pEntries(E) = RGB(B, G, R)
                
            Next
            ucAniGIF(2).SetPalette P, pEntries()
        End If
    Next
    
End Sub

Private Sub chkSolidBkg_Click()
    If chkSolidBkg = vbChecked Then
        ucAniGIF(1).BackColor = vbCyan
        ucAniGIF(1).BackStyle = gfbSolid
    Else
        ucAniGIF(1).BackStyle = gfbTransparent
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim I As Integer, Action As AnimationActions
    Select Case Index
        Case 0: Action = gfaStop
        Case 1: Action = gfaPause
        Case 2: Action = gfaForward
        Case 3: Action = gfaPlay
    End Select
    For I = ucAniGIF.LBound To ucAniGIF.UBound
        ucAniGIF(I).Action = Action
    Next
End Sub

Private Sub chkMirror_Click()
    Dim I As Integer
    For I = ucAniGIF.LBound To ucAniGIF.UBound  ' toggle horizontal mirroring
        ucAniGIF(I).Mirrored = ucAniGIF(I).Mirrored Xor gfmHorizontal
    Next
End Sub

Private Sub Command2_Click()

    If MsgBox("Two things." & vbCrLf & _
        "1. Click Yes to retrieve the GIF, no links will be activated." & vbCrLf & _
        "2. Notice the images disappearing? Once OCX is compiled or EXE is compiled, this won't happen", _
        vbInformation + vbYesNo, "Continue?") = vbNo Then Exit Sub

    ' this will not activate the link
    ucAniGIF(3).LoadAnimatedGIF_Remote "http://www.animation-station.com/shared/coollinks.gif"
    ' the remote path can be a server too: i.e., \\myserver\my pictures\mygif.gif or a hard drive path
End Sub

Private Sub Command3_Click()
    Form2.Show
End Sub

Private Sub Form_Load()
    chkMirror.BackColor = Me.BackColor
    chkModPal.BackColor = Me.BackColor
    chkModClrTbl2.BackColor = Me.BackColor
    chkSolidBkg.BackColor = Me.BackColor
    Command2.ToolTipText = "Target URL is http://www.animation-station.com/shared/coollinks.gif"
End Sub

Private Sub Form_Resize()
    ' Example of pausing animation while minimized
    Dim I As Integer
    Dim oldAction As AnimationActions, newAction As AnimationActions
    
    If Me.WindowState = vbMinimized Then
        oldAction = gfaPlay
        newAction = gfaPause
    Else
        oldAction = gfaPause
        newAction = gfaPlay
    End If
    
    For I = ucAniGIF.LBound To ucAniGIF.UBound
        If ucAniGIF(I).Action = oldAction Then ucAniGIF(I).Action = newAction
    Next

End Sub

Private Sub ucAniGIF_FrameChanged(Index As Integer, ByVal FrameIndex As Long, viaTimer As Boolean)
    ' event calls back each time a frame is rendered, should you want this info
    If Index = 2 Then lblFrame(Index).Caption = "Frame " & FrameIndex
    ' to see the frame change indication for all the example GIFS, unrem below and rem above
    'lblFrame(Index).Caption = "Frame " & FrameIndex
End Sub

Private Sub ucAniGIF_LoopsEnded(Index As Integer)
    ' should you want to know when a GIF terminates its
    ' loop and stops animating. You will also get this
    ' event for a single frame GIF each time it is displayed.
    ucAniGIF(Index).Action = gfaReset ' simply restart
End Sub

Private Sub ucAniGIF_RemoteLoadComplete(Index As Integer, ByVal gifWidth As Single, ByVal gifHeight As Single, ByRef Cancel As Boolean)
    ' When you called the LoadAnimatedGIF_Remote routine, (See Command2_Click)
    ' this event will be fired if the file was successfully read and the header
    ' of the file indicates it is a GIF.
    
    ' Set Cancel to True to prevent loading it.
    ' Otherwise, it will be displayed with the current settings of the usercontrol
    Command2.Enabled = False  ' successfully downloaded sample URL gif
    
    With ucAniGIF(Index)                ' to be a little thorough
        Set .AnimatedGIF = Nothing  ' remove previous image before changing attributes
        .Stretch = gfsShrinkScaleToFit ' set scale
        .DelayAnimation = gfdNone ' set delay mode
        .Mirrored = gfmNone         ' set mirror options
        .Enabled = True             ' enable it
    End With                        ' next, the image will be processed and displayed

End Sub

Private Sub ucAniGIF_RemoteLoadFailure(Index As Integer)
    ' When you called the LoadAnimatedGIF_Remote routine, this event will be fired if the
    ' file was NOT successfully read, errors occurred or the header of the file indicates
    ' it is NOT a GIF.
    MsgBox "Failed to download/read the remote GIF file. Possible server is down or " & vbCrLf & _
        "the GIF no longer exists. To test this functionality..." & vbCrLf & _
        "1. Go to any website that is displaying a GIF" & vbCrLf & _
        "2. Right click on the GIF and select Properties from the menu" & vbCrLf & _
        "3. Highlight the complete URL and copy it: Right click, copy" & vbCrLf & _
        "4. Paste the URL into the Command2_Click event. Try again.", vbInformation + vbOKOnly
    Command2.Enabled = False
End Sub
