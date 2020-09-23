VERSION 5.00
Object = "*\A..\..\..\..\..\..\MYDOCU~1\WORKSH~1\VB\DANCIN~1\ANIMAR~1\prjAnimatedLabelC.vbp"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjAnimatedLabelC.AnimatedLabelC AnimatedLabelC1 
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1720
      Picture         =   "frmTest1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AnimatedLabelC1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  End
End Sub

Private Sub Form_Load()
  Dim msg As String

  AnimatedLabelC1.Move 0, 0, Me.Width, Me.Height
  AnimatedLabelC1.ForeColor = vbYellow
  AnimatedLabelC1.StartupYpos = 150
  AnimatedLabelC1.StartupWait = 1.5         'Waits 1.5 sec at startup
  AnimatedLabelC1.ScrollStart = 300
  AnimatedLabelC1.ScrollEnd = -850
  AnimatedLabelC1.SkipPixels = 1
'Available Tags List:>
'<F.BOLD>,<F.ITALIC>,<F.STRIKETHROUGH>,<F.UNDERLINE>
'<F.WGHT:,<F.SIZEX:,<F.CAHRS:,<F.COLOR:,<F.NAMEX: <eg:<F.COLOR:&H000000FF>
'The tags are not case sensitive
'See Script.txt for more informations
'If you load scripts from file make the script file hidden,readonly and extension like .exe/.dll/.sys to rescue from unautherised editing.
'You can load information from file by giving caption = afilename
'eg : Set AnimatedLabelC1.Caption = "atextfile.sys"

  msg = "Animated Label Style C<F.BOLD>;"   'Displays in bold
  msg = msg & "Activex Control<F.COLOR:&H00FF00FF&>/b"           'Another way to display in bold
  msg = msg & ";/l;"                        'Displays a doted line
  msg = msg & ";;;;;;;"                     'Displays 7 blank lines
  msg = msg & "Version<F.COLOR:&H0000FF00&>/i/b;"                'Italic
  msg = msg & "1.0.0;"
  msg = msg & ";;;;;;;"
  msg = msg & "A Program By<F.NAMEX:Comic Sans MS>/b;"
  msg = msg & "Satheesh Chandran<F.NAMEX:Comic Sans MS>;"
  msg = msg & ";;;;;;;"
  msg = msg & "Copyright<F.COLOR:&H0000FF00&>;"
  msg = msg & "(C) 2004 - 2005 Satheesh Chandran;"
  msg = msg & ";;;;;;;"
  msg = msg & "Contact/i/b;"
  msg = msg & "vigyanabikshu@hotmail.com/i;"
  msg = msg & ";;;;;;;"
  msg = msg & "Where is Truth There is Success;/l;/l;/l;"

  AnimatedLabelC1.Caption = msg
  Set AnimatedLabelC1.Picture = LoadPicture(App.Path & "\Flower.bmp")

  Me.Caption = "Scroll From " & AnimatedLabelC1.ScrollStart & "  To " & AnimatedLabelC1.ScrollEnd
  
  
  Load Form2
  Form2.Show
  Form2.Move 0, 0
End Sub
