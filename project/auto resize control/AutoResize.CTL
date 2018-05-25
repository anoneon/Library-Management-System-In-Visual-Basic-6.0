VERSION 5.00
Begin VB.UserControl AutoResize 
   CanGetFocus     =   0   'False
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   750
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   Picture         =   "AutoResize.ctx":0000
   ScaleHeight     =   720
   ScaleWidth      =   750
   ToolboxBitmap   =   "AutoResize.ctx":2132
End
Attribute VB_Name = "AutoResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type ControlSettings
    IndexValue As Long
    LeftValue As Single
    TopValue As Single
    WidthValue As Single
    HeightValue As Single
    FontSizeValue As Single
End Type
Private State As Boolean, FontResize As Boolean, AspectRatio43 As Boolean, HM As Long, WM As Long, AspectRatioValue43 As Double, CtrlTot As Long
Private ControlsOnForm() As ControlSettings, ParentWidth As Single, ParentHeight As Single
Private WithEvents ParentForm As Form
Attribute ParentForm.VB_VarHelpID = -1
Private Sub ParentForm_Load()
CtrlTot = 0
Call StoreOriginalSettings
End Sub
Private Sub UserControl_InitProperties()
'Set Default properties values
FontResizable = True 'Resizing font enabled
KeepAspectRatio = False 'Keeping aspect ration disabled
HMin = 3000 'Set the minimum form Height allowed
WMin = 4800 'Set the minimum form Width allowed
AspectRatioValue = 0 'You can set it manually or let the program to do it by enabling KeepAspectRatio property
Enabled = True  'Autoresize Control enabled
UserControl.Extender.Tag = "NO" 'Don't include this control among the ones to resize
'UserControl.Extender.Name = "Resize"    'Name to assign to Autoresize control when you put it on a form
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
FontResizable = PropBag.ReadProperty("FontResizable", True)
KeepAspectRatio = PropBag.ReadProperty("KeepAspectRatio", False)
AspectRatioValue = PropBag.ReadProperty("AspectRatioValue", 0)
Enabled = PropBag.ReadProperty("Enabled", True)
HMin = PropBag.ReadProperty("HMin", 3000)
WMin = PropBag.ReadProperty("WMin", 4800)
If Ambient.UserMode = False Then Exit Sub 'If not running exit
Set ParentForm = UserControl.Parent 'ParentForm is the form containing the autoresize control
End Sub
Private Sub UserControl_Resize()
UserControl.Width = 400 'Width of autoresize control
UserControl.Height = 400 'Height of autoresize control
End Sub
Private Sub ParentForm_Resize()
'Triggered when the form containing the autoresize control is resized
If ParentForm.WindowState = 1 Then Exit Sub 'If reduced to icon dont't resize
If Not Enabled Then Exit Sub    'If the autoresize control is disabled don't resize
If ParentForm.Height < HMin Then ParentForm.Height = HMin   'If form height < than value assigned to HMIN force the parent height to HMIN
If ParentForm.Width < WMin Then ParentForm.Width = WMin 'same as above for widtyh
Call UpdateControls 'Resize the controls on the form
End Sub
Private Sub StoreOriginalSettings()
'This routine runs just once (when the form is loaded)
On Error Resume Next
Dim i As Integer, Controllo As Control
If Not Ambient.UserMode Then Exit Sub 'If not running exit
ParentWidth = ParentForm.ScaleWidth 'save the original size of the form
ParentHeight = ParentForm.ScaleHeight
For i = 0 To ParentForm.Controls.Count - 1  'Loop to check all the controls on the form
    Set Controllo = ParentForm.Controls(i)
    If TypeName(Controllo) = "StatusBar" Then 'if the control is a status bar
        If Controllo.Align > 0 Then GoTo Skip 'it will be resized only if the Align property is set to none
        'You can add here other controls that have similar properties
    End If
    If Controllo.Tag = "NO" Then GoTo Skip 'The Tag property set to NO prevent the control from resizing
    CtrlTot = CtrlTot + 1   'Variable to keep the number of controls to resize
    ReDim Preserve ControlsOnForm(1 To CtrlTot) 'Redim the array containing the data needed for resizing
    With ControlsOnForm(CtrlTot)
        .IndexValue = i 'save the index of the control
        If TypeName(Controllo) = "Line" Then    'the line control doesn't have Left, Top.... properties, so you must manage it differently
            'you can add here other particular controls - the ones that don't have standard properties
            .LeftValue = Controllo.X1 'save the original size
            .TopValue = Controllo.Y1
            .WidthValue = Controllo.X2
            .HeightValue = Controllo.Y2
        Else
            If Controllo.Left < 0 And TypeName(Controllo.Container) = "SSTab" Then
                .LeftValue = Controllo.Left + 75000 'add 75000 to record the correct position of controls not located in the current tab
            Else
                .LeftValue = Controllo.Left 'save the original size
            End If
            .TopValue = Controllo.Top
            .WidthValue = Controllo.Width
            .HeightValue = Controllo.Height
            .FontSizeValue = Controllo.Font.Size
        End If
    End With
Skip:
Next i
End Sub
Private Sub UpdateControls()
On Error Resume Next
Dim i As Integer, FFactor As Single, WFactor As Single, HFactor As Single
Static ChangingRatio As Boolean
If Not Ambient.UserMode Then Exit Sub 'if not running exit
If ChangingRatio Then Exit Sub 'prevent recursive calls if KeepAspectRatio is True
If KeepAspectRatio And AspectRatioValue > 0 And ParentForm.WindowState = 0 Then 'if the form is not icon or maximized
    ChangingRatio = True
    ParentForm.Height = AspectRatioValue * ParentForm.Width 'change the form height to keep aspect ratio
    ChangingRatio = False
End If
WFactor = ParentForm.ScaleWidth / ParentWidth 'calculates the increasing or decreasing factor to use
HFactor = ParentForm.ScaleHeight / ParentHeight
'set the font increasing or decreasing factor to the minimum width-height factor
If WFactor < HFactor Then
  FFactor = WFactor
Else
  FFactor = HFactor
End If
For i = 1 To CtrlTot 'loop through the controls included in the ControlsOnForm array
    With ControlsOnForm(i)
        If TypeName(ParentForm.Controls(.IndexValue)) = "Line" Then 'if it's a line manage it differently
            'if you added other controls in the StoreOriginalSettings routine
            'add the same controls here
            ParentForm.Controls(.IndexValue).X1 = .LeftValue * WFactor
            ParentForm.Controls(.IndexValue).Y1 = .TopValue * HFactor
            ParentForm.Controls(.IndexValue).X2 = .WidthValue * WFactor
            ParentForm.Controls(.IndexValue).Y2 = .HeightValue * HFactor
        Else
            If FontResizable Then ParentForm.Controls(.IndexValue).Font.Size = .FontSizeValue * FFactor 'resize the fonts if you enabled Font resizing
            If ParentForm.Controls(.IndexValue).Left < 0 And TypeName(ParentForm.Controls(.IndexValue).Container) = "SSTab" Then
                ParentForm.Controls(.IndexValue).Left = .LeftValue * WFactor - 75000 'subtract 75000 to keep controls hidden (the ones not located in the current tab)
            Else
                ParentForm.Controls(.IndexValue).Left = .LeftValue * WFactor 'resize the control multiplying the original size for the calculated factor
            End If
            ParentForm.Controls(.IndexValue).Top = .TopValue * HFactor
            ParentForm.Controls(.IndexValue).Width = .WidthValue * WFactor
            ParentForm.Controls(.IndexValue).Height = .HeightValue * HFactor
        End If
    End With
Next i
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'Save the values assigned to properties
Call PropBag.WriteProperty("FontResizable", FontResizable, True)
Call PropBag.WriteProperty("KeepAspectRatio", KeepAspectRatio, False)
Call PropBag.WriteProperty("AspectRatioValue", AspectRatioValue)
Call PropBag.WriteProperty("HMin", HMin, 3000)
Call PropBag.WriteProperty("WMin", WMin, 4800)
Call PropBag.WriteProperty("Enabled", Enabled, True)
End Sub
Public Property Get Enabled() As Boolean
Enabled = State
End Property
Public Property Let Enabled(ByVal Value As Boolean)
State = Value
End Property
Public Property Get FontResizable() As Boolean
FontResizable = FontResize
End Property
Public Property Let FontResizable(ByVal Value As Boolean)
FontResize = Value
End Property
Public Property Get KeepAspectRatio() As Boolean
KeepAspectRatio = AspectRatio43
End Property
Public Property Let KeepAspectRatio(ByVal Value As Boolean)
On Error Resume Next
AspectRatio43 = Value
'if AspectRation enabled, calculates the AspectRatio Value
If Value Then AspectRatioValue = UserControl.Extender.Parent.Height / UserControl.Extender.Parent.Width Else AspectRatioValue = 0
PropertyChanged "KeepAspectRatio"
End Property
Public Property Get AspectRatioValue() As Double
AspectRatioValue = AspectRatioValue43
End Property
Public Property Let AspectRatioValue(ByVal Value As Double)
AspectRatioValue43 = Value
End Property
Public Property Get HMin() As Long
HMin = HM
End Property
Public Property Let HMin(ByVal Value As Long)
HM = Value
End Property
Public Property Get WMin() As Long
WMin = WM
End Property
Public Property Let WMin(ByVal Value As Long)
WM = Value
End Property

