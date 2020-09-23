VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   3000
      TabIndex        =   2
      Top             =   1860
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   2340
      TabIndex        =   1
      Top             =   540
      Width           =   855
   End
   Begin Project1.ctlGlobe UserControl11 
      Height          =   1575
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      _extentx        =   3201
      _extenty        =   2778
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.UserControl11.Start

End Sub

Private Sub Command2_Click()
Me.UserControl11.Off
End Sub
