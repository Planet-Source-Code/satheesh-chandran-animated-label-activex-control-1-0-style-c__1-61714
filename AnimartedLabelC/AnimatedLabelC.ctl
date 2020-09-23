VERSION 5.00
Begin VB.UserControl AnimatedLabelC 
   AutoRedraw      =   -1  'True
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   ScaleHeight     =   64
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   254
   Begin VB.Timer tmrRefresh 
      Interval        =   10
      Left            =   480
      Top             =   480
   End
End
Attribute VB_Name = "AnimatedLabelC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************
'Dancing Animated Label Style C Actiex Control
'Your Suggessions and comments feel free to contact me at
'vigyana_bikshu@yahoo.com
'Licence : GNU General Public Licence
'Feel free to use these codes in your projects.
'Satheesh Chandran
'Date : 7-7-2005
'*********************************************************


Option Explicit


'Enums
Public Enum ENUM_BorderStyle
  None
  FixedSingle
End Enum

Public Enum ENUM_Alignment
  LeftJustify
  Center
End Enum

'Public Enum ENUM_CaptionStyle
  'Normal            'user selected font for each line
  'ScriptEnabled     'User selected font will avoid and take all font properties from the Script in the caption
'  FromFile          'User selected font will avoid and take all font properties from the Script and will read from specified script file
'End Enum

 
'Types

Private Type MyFont
  Bold As Boolean
  Charset As Long
  Italic As Boolean
  Name As String
  Size As Long
  Strikethrough As Boolean
  Underline As Boolean
  Weight As Long
End Type

'A line of text with Font style attributes
Private Type aLine
  line As String
  Font As MyFont
  ForeColor As Long
End Type

'Constants
'private Const DFLT_Caption As String = Extender.Name
Private Const DFLT_Scroll As Boolean = True
Private Const DFLT_borderstyle As Byte = 0
Private Const DFLT_SkipPixels = 1
Private Const DFLT_ScrollStart As Long = 50
Private Const DFLT_ScrollEnd As Long = -50
Private Const DFLT_Alignment As Byte = 1        'Center
Private Const DFLT_StartupWait As Double = 0.01


'Variables
Private TEST_MODE As Boolean
Private DFLT_ScriptFile As String
Private DFLT_BackColor As Long
Private DFLT_ForeColor As Long
Private DFLT_TimerInterval As Long
Private DFLT_StartupYpos As Long
Private PropertiesReady As Boolean
Private iCaption As String
Private iForeColor As Long
Private iSkipPixels As Integer
Private iScrollStart As Long
Private iScrollEnd As Long
Private iScriptFile As String
Private iAlignment As ENUM_Alignment
Private iStartupWait As Double
Private iStartupYpos As Long
Private displayBuffer(1 To 500) As aLine
Private lineCount As Integer
Private theFont As IFont
Private icurrentYpos As Long


'Events
'******
Public Event click()
Public Event Dblclick()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Public Event MouseEnter(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Public Event MouseExit()
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event Paint()

Private Sub Usercontrol_Click()
  RaiseEvent click
End Sub
Private Sub Usercontrol_DblClick()
  RaiseEvent Dblclick
End Sub
Private Sub Usercontrol_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub
Private Sub Usercontrol_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub Usercontrol_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub Usercontrol_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Usercontrol_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub
Private Sub Usercontrol_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub
Private Sub Usercontrol_Paint()
    RaiseEvent Paint
End Sub


'Properties
Public Property Get TimerInterval() As Long
  TimerInterval = tmrRefresh.Interval
End Property
Public Property Let TimerInterval(ByVal vNewValue As Long)
  tmrRefresh.Interval = vNewValue
End Property

Public Property Get Caption() As String
  Caption = iCaption
End Property
Public Property Let Caption(ByVal vNewValue As String)
  iCaption = vNewValue
  Call SetupMainBufferLines
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
  UserControl.BackColor = vNewValue
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = iForeColor
End Property
Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
   iForeColor = vNewValue
   Call SetupMainBufferLines
End Property

Public Property Get Scroll() As Boolean
  Scroll = tmrRefresh.Enabled
End Property
Public Property Let Scroll(ByVal vNewValue As Boolean)
  tmrRefresh.Enabled = vNewValue
End Property

Public Property Get Picture() As Picture
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal newPicture As Picture)
  Set UserControl.Picture = newPicture
End Property

Public Property Get BorderStyle() As ENUM_BorderStyle
  BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal vNewValue As ENUM_BorderStyle)
  UserControl.BorderStyle = vNewValue
End Property

Public Property Get SkipPixels() As Integer
  SkipPixels = iSkipPixels
End Property
Public Property Let SkipPixels(ByVal vNewValue As Integer)
  iSkipPixels = vNewValue
End Property

Public Property Get ScrollStart() As Long
  ScrollStart = iScrollStart
End Property
Public Property Let ScrollStart(ByVal vNewValue As Long)
  iScrollStart = vNewValue
  DFLT_StartupYpos = vNewValue
End Property

Public Property Get ScrollEnd() As Long
  ScrollEnd = iScrollEnd
End Property
Public Property Let ScrollEnd(ByVal vNewValue As Long)
  iScrollEnd = vNewValue
End Property

Public Property Get Alignment() As ENUM_Alignment
  Alignment = iAlignment
End Property
Public Property Let Alignment(ByVal vNewValue As ENUM_Alignment)
  iAlignment = vNewValue
End Property

Public Property Get StartupWait() As Double
  StartupWait = iStartupWait
End Property
Public Property Let StartupWait(ByVal vNewValue As Double)
  iStartupWait = vNewValue
End Property

Public Property Get StartupYpos() As Long
  StartupYpos = iStartupYpos
End Property
Public Property Let StartupYpos(ByVal vNewValue As Long)
  iStartupYpos = vNewValue
End Property

'Public Property Get Font() As IFont
'   Set Font = theFont
'End Property
'Public Property Let Font(newFont As IFont)
'   pSetFont newFont
'End Property
'Public Property Set Font(iFnt As IFont)
'   pSetFont iFnt
'End Property
'Private Sub pSetFont(iFnt As IFont)
'  Set theFont = iFnt
'  PropertyChanged "Font"
'  Call SetupMainBufferLines
'End Sub

Public Property Get CurrentYpos() As Long
  CurrentYpos = icurrentYpos
End Property
Public Property Let CurrentYpos(ByVal vNewValue As Long)
  CurrentYpos = vNewValue
End Property



'Code
Private Sub tmrRefresh_Timer()
If TEST_MODE = False Then On Error Resume Next
  Static TMP As Boolean, i As Long
  UserControl.Cls
  If TMP = False Then       'If first time then wait 1 second
    TMP = True
    fillCaption iStartupYpos        'displays caption from ith ypos
    i = iStartupYpos
    SlowDown iStartupWait
  End If
  fillCaption i        'displays caption from ith ypos
  If i > iScrollEnd Then        'reset ypos
    i = i - SkipPixels
  Else
    i = iScrollStart
  End If
  
End Sub

'Just fills the usercontrol with the caption with given format
Private Sub fillCaption(ypos As Long)
If TEST_MODE = False Then On Error Resume Next
  UserControl.CurrentY = ypos
  icurrentYpos = ypos
  Dim i As Integer
  For i = 1 To lineCount
    With UserControl.Font
        .Bold = displayBuffer(i).Font.Bold
        .Charset = displayBuffer(i).Font.Charset
        .Italic = displayBuffer(i).Font.Italic
        .Name = displayBuffer(i).Font.Name
        If displayBuffer(i).Font.Size > 0 Then
          .Size = displayBuffer(i).Font.Size
        End If
        .Strikethrough = displayBuffer(i).Font.Strikethrough
        .Underline = displayBuffer(i).Font.Underline
        '.Weight = displayBuffer(i).Font.Weight
    End With
    'Set the location of the text
    UserControl.CurrentY = UserControl.CurrentY + 2 'Line Spacing
    If iAlignment = Center Then UserControl.CurrentX = (UserControl.ScaleWidth - UserControl.TextWidth(displayBuffer(i).line)) / 2 'Center Alignment
    UserControl.ForeColor = displayBuffer(i).ForeColor
    UserControl.Print displayBuffer(i).line
  Next i
  Exit Sub
'err:
End Sub


Private Sub UserControl_Initialize()
TEST_MODE = True    'Make this false while making ocx to rescue from errors if any.
If TEST_MODE = False Then On Error Resume Next
  DFLT_BackColor = UserControl.BackColor
  DFLT_ForeColor = UserControl.ForeColor
  DFLT_TimerInterval = tmrRefresh.Interval
  'Set DFLT_Font = UserControl.Font
End Sub

Private Sub UserControl_InitProperties()
If TEST_MODE = False Then On Error Resume Next
  Caption = Extender.Name
  TimerInterval = DFLT_TimerInterval
  BackColor = DFLT_BackColor
  ForeColor = DFLT_ForeColor
  Scroll = DFLT_Scroll
  BorderStyle = DFLT_borderstyle
  SkipPixels = DFLT_SkipPixels
  ScrollStart = DFLT_ScrollStart
  ScrollEnd = DFLT_ScrollEnd
  'ScriptFile = DFLT_ScriptFile
  Alignment = DFLT_Alignment
  StartupWait = DFLT_StartupWait
  StartupYpos = DFLT_StartupYpos
'  Font = DFLT_Font
  MsgBox "Remember to set proper values to 'ScrollStart' and 'ScrollEnd' properties." & Chr(13) & "If the caption not displayed first time don't worry, It will display when you run.", vbInformation
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  On Error Resume Next      'Requires the line On Error Resume Next
  Caption = PropBag.ReadProperty("Caption", Extender.Name)
  TimerInterval = PropBag.ReadProperty("TimerInterval", DFLT_TimerInterval)
  BackColor = PropBag.ReadProperty("BackColor", DFLT_BackColor)
  ForeColor = PropBag.ReadProperty("ForeColor", DFLT_ForeColor)
  Scroll = PropBag.ReadProperty("Scroll", DFLT_Scroll)
  Set Picture = PropBag.ReadProperty("Picture")
  BorderStyle = PropBag.ReadProperty("borderstyle", DFLT_borderstyle)
  SkipPixels = PropBag.ReadProperty("SkipPixels", DFLT_SkipPixels)
  ScrollStart = PropBag.ReadProperty("ScrollStart", DFLT_ScrollStart)
  ScrollEnd = PropBag.ReadProperty("ScrollEnd", DFLT_ScrollEnd)
  'ScriptFile = PropBag.ReadProperty("ScriptFile", DFLT_ScriptFile)
  Alignment = PropBag.ReadProperty("Alignment", DFLT_Alignment)
  StartupWait = PropBag.ReadProperty("StartupWait", DFLT_StartupWait)
  StartupYpos = PropBag.ReadProperty("StartupYpos", DFLT_StartupYpos)
  'Font = PropBag.ReadProperty("Font", DFLT_Font)
  'CaptionStyle = PropBag.ReadProperty("CaptionStyle", DFLT_CaptionStyle)
  PropertiesReady = True
  Call SetupMainBufferLines
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
If TEST_MODE = False Then On Error Resume Next
  PropBag.WriteProperty "Caption", Caption, Extender.Name
  PropBag.WriteProperty "TimerInterval", TimerInterval, DFLT_TimerInterval
  PropBag.WriteProperty "BackColor", BackColor, DFLT_BackColor
  PropBag.WriteProperty "ForeColor", ForeColor, DFLT_ForeColor
  PropBag.WriteProperty "Scroll", Scroll, DFLT_Scroll
  PropBag.WriteProperty "Picture", Picture
  PropBag.WriteProperty "borderstyle", BorderStyle, DFLT_borderstyle
  PropBag.WriteProperty "SkipPixels", SkipPixels, DFLT_SkipPixels
  PropBag.WriteProperty "ScrollStart", ScrollStart, DFLT_ScrollStart
  PropBag.WriteProperty "ScrollEnd", ScrollEnd, DFLT_ScrollEnd
  'PropBag.WriteProperty "ScriptFile", ScriptFile, DFLT_ScriptFile
  PropBag.WriteProperty "Alignment", Alignment, DFLT_Alignment
  PropBag.WriteProperty "StartupWait", StartupWait, DFLT_StartupWait
  PropBag.WriteProperty "StartupYpos", StartupYpos, DFLT_StartupYpos
  'PropBag.WriteProperty "Font", Font, DFLT_Font
  'PropBag.WriteProperty "CaptionStyle", CaptionStyle, DFLT_CaptionStyle
End Sub














'Other Functions

'This function will fill the mainBuffer() that the texts to be displayed
Public Sub SetupMainBufferLines()
  If TEST_MODE = False Then On Error Resume Next
  
  Dim tmpFnt As MyFont
  Dim tmpForeColor As Long

  If iCaption = "" Then Exit Sub
  If PropertiesReady = False Then Exit Sub
  
  Dim rawCaption()  As String
  'If the caption is a valid file name then select the contents of that file as caption
  Dim fileContent As String
  If FileExist(iCaption) = True Then    'Check the caption is a file name specification ?
    fileContent = getFileContent(iCaption)  'Then loads the contents of that file
    rawCaption = Split(fileContent, ";")
  ElseIf FileExist(App.Path & "\" & iCaption) Then  'If the caption is a file specification without path information then check for the file on app.path
    fileContent = getFileContent(App.Path & "\" & iCaption)
    rawCaption = Split(fileContent, ";")
  Else
    rawCaption = Split(iCaption, ";")   'Splits the caption with ";". The ";" will consider as a new line
  End If

  Erase displayBuffer                 'Erasese the old values if any
  lineCount = 0
  Dim i As Integer
  For i = 0 To UBound(rawCaption)
      tmpFnt = getDefaultMyFont()    'Initialises user selected font
      tmpForeColor = iForeColor     'Initializes user selected forecolor
      If rawCaption(i) = "" Then
        addLine tmpFnt        'Adds a blank new line
      ElseIf Left(rawCaption(i), 1) = "<" Then      '< is comment entry
        
      Else
        Dim t As Integer
        
        t = InStr(1, UCase(rawCaption(i)), "<F.BOLD>")  'If the line contains "<F.BOLD>" then set bold property to true
        If t > 0 Then tmpFnt.Bold = True
        
        t = InStr(1, UCase(rawCaption(i)), "<F.ITALIC>")
        If t > 0 Then tmpFnt.Italic = True
        
        t = InStr(1, UCase(rawCaption(i)), "<F.STRIKETHROUGH>")
        If t > 0 Then tmpFnt.Strikethrough = True
        
        t = InStr(1, UCase(rawCaption(i)), "<F.UNDERLINE>")
        If t > 0 Then tmpFnt.Underline = True
          
        t = InStr(1, UCase(rawCaption(i)), "<F.WGHT:")
        If t > 0 Then tmpFnt.Weight = Val(getValue(rawCaption(i), t))
          
        t = InStr(1, UCase(rawCaption(i)), "<F.SIZEX:")
        If t > 0 Then tmpFnt.Size = Val(getValue(rawCaption(i), t))
          
        t = InStr(1, UCase(rawCaption(i)), "<F.CAHRS:")
        If t > 0 Then tmpFnt.Charset = Val(getValue(rawCaption(i), t))
        
        t = InStr(1, UCase(rawCaption(i)), "<F.NAMEX:")
        If t > 0 Then tmpFnt.Name = getValue(rawCaption(i), t)
       
        'For the tags /B,-B at the end of a line or /L
        If UCase(Right(rawCaption(i), 2)) = "/B" Then
          tmpFnt.Bold = True
          rawCaption(i) = Left(rawCaption(i), Len(rawCaption(i)) - 2)
        End If
        If UCase(Right(rawCaption(i), 2)) = "/I" Then   'Italic
          tmpFnt.Italic = True
          rawCaption(i) = Left(rawCaption(i), Len(rawCaption(i)) - 2)
        End If
        If UCase(Right(rawCaption(i), 2)) = "/L" Then
           rawCaption(i) = Left(rawCaption(i), Len(rawCaption(i)) - 2)
           rawCaption(i) = String$(50, "-")
        End If

        t = InStr(1, UCase(rawCaption(i)), "<F.COLOR:")
        If t > 0 Then tmpForeColor = Val(getValue(rawCaption(i), t))
        
        'Adds the line to the mainBuffer()
        addLine tmpFnt, tmpForeColor, removeTags(rawCaption(i))
      End If
  Next i
End Sub
'Removes tags like <F.Bold>, <F.COLOR> from the line
Private Function removeTags(ByVal str As String) As String
If TEST_MODE = False Then On Error Resume Next
  If str = "" Then Exit Function
  Dim i As Integer
  For i = 1 To Len(str)
    If Mid(str, i, 1) = "<" Then
      removeTags = Left(str, i - 1)  'all the properties are enclosed with < >, here removes texts with in < >
      Exit Function
    End If
  Next i
  removeTags = str  'if not found any <
End Function
'Returns the values from the tags that contains numbers like  <F.COLOR:56548>, here 56548 must return
Private Function getValue(ByVal str As String, pos As Integer) As String
If TEST_MODE = False Then On Error Resume Next
  If str = "" Then Exit Function
  Dim i As Integer, t As String
  For i = pos + 9 To Len(str)
    t = Mid(str, i, 1)
    If t <> ">" Then
      getValue = getValue & Mid(str, i, 1)
    Else
      Exit Function
    End If
  Next i
End Function

Private Sub addLine(myFnt As MyFont, Optional ByVal theForeColor As Long, Optional theText As String = "")
If TEST_MODE = False Then On Error Resume Next
    lineCount = lineCount + 1
    displayBuffer(lineCount).line = theText   'adds the line to the buffer
    displayBuffer(lineCount).ForeColor = theForeColor
    displayBuffer(lineCount).Font = myFnt
    'MsgBox myFnt.Bold
End Sub

Private Sub SlowDown(itime As Double)
  'This subroutine will simply wait a number of seconds
  ' for example you can place the code slowdown 1 in your app
  'to make it wait 1 second. It will only recognize integer values
  Dim lStartTime As Double
  lStartTime = Timer
  Do While Not Timer >= lStartTime + itime
    DoEvents
  Loop
End Sub

'Check for a file exists or not
Private Function FileExist(fileName As String) As Boolean
  On Error GoTo err
  If Trim(fileName) = "" Then
    FileExist = False
    Exit Function
  End If
 
  FileExist = IIf(Dir(fileName) <> "", True, False)

Exit Function
err:
End Function

Private Function getFileContent(ByVal theFile As String) As String
  On Error GoTo err
  Dim fileNum As Integer, aLine As String
  fileNum = FreeFile
  Open theFile For Input As #fileNum
  While Not EOF(fileNum)
    Line Input #fileNum, aLine
    getFileContent = getFileContent & aLine & ";"
  Wend
  Close #fileNum
  Exit Function
err:
  If err.Number = 53 Then
    MsgBox "File " & theFile & Chr(13) & "Not Found", vbCritical
  End If
End Function

Private Function getDefaultMyFont() As MyFont
  With getDefaultMyFont
    .Bold = False
    .Charset = 0
    .Italic = False
    .Name = "MS Sans Serif"
    .Size = 10
    .Strikethrough = False
    .Underline = False
    .Weight = 10
  End With
End Function
