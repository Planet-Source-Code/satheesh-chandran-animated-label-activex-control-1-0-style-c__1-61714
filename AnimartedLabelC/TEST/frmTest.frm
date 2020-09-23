VERSION 5.00
Object = "*\A..\..\..\..\..\..\MYDOCU~1\WORKSH~1\VB\DANCIN~1\ANIMAR~1\prjAnimatedLabelC.vbp"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4695
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjAnimatedLabelC.AnimatedLabelC AnimatedLabelC2 
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   3240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1085
      Picture         =   "frmTest.frx":0000
   End
   Begin prjAnimatedLabelC.AnimatedLabelC AnimatedLabelC1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5741
      Caption         =   "script.txt"
      BackColor       =   0
      ForeColor       =   65535
      Picture         =   "frmTest.frx":001C
      ScrollStart     =   150
      ScrollEnd       =   -900
      StartupYpos     =   50
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AnimatedLabelC1_click()
  If AnimatedLabelC1.Scroll = True Then
    AnimatedLabelC1.Scroll = False
  Else
    AnimatedLabelC1.Scroll = True
  End If
End Sub

Private Sub Form_Load()
  Me.Caption = App.Path & "\" & AnimatedLabelC1.Caption
  AnimatedLabelC1.ToolTipText = Me.Caption
End Sub
