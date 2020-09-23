VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4095
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrHighlight 
      Interval        =   1000
      Left            =   7275
      Top             =   3750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Oke"
      Height          =   405
      Left            =   3263
      TabIndex        =   5
      Top             =   3690
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   -75
      Width           =   7875
      Begin VB.TextBox Text1 
         Height          =   2970
         Left            =   75
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "frmSplash.frx":08CA
         Top             =   720
         Width           =   5760
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "by: JosGroen@hotmail.com"
         Height          =   360
         Left            =   5850
         TabIndex        =   4
         Top             =   615
         Width           =   2010
      End
      Begin VB.Image imgLogo 
         Height          =   2550
         Left            =   5850
         Picture         =   "frmSplash.frx":08D2
         Top             =   870
         Width           =   2010
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5775
         TabIndex        =   1
         Top             =   3345
         Width           =   1995
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   45
         TabIndex        =   2
         Top             =   135
         Width           =   6990
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & App.Revision
    lblProductName.Caption = App.Title
    Me.Text1.Text = "Planet Source Code File Fetcher " & App.Major & "." & App.Minor & App.Revision & vbCrLf & _
"Convert your daily planet source code mail to a database" & vbCrLf & _
"V1.16 :" & "Add: Select your own Logfile (Settings)" & vbCrLf & _
"V1.16 :" & "Fix: Search and enter" & vbCrLf & _
"V1.16 :" & "Fix: Listview one title" & vbCrLf & _
"V1.16 :" & "Fix: ModStart" & vbCrLf & _
"V1.16 :" & "Fix: DBGrid32.ocx removed reference (Thanx Wendell Jackson)" & vbCrLf & _
"V1.16 :" & "Add: Hyperlink to Author homepage (if any)" & vbCrLf & "" & vbCrLf & _
"New Download the 'Copy and past friendly page'  if there is no zip file" & vbCrLf & _
"NEW in this version , Tree view object for easy browsing" & vbCrLf & _
"NEW author will be get from page (if any) and store in the tblAU_author table" & vbCrLf & _
"NEW Search function" & vbCrLf & _
"NEW GUI  (for Chris Rose)" & vbCrLf & _
"" & vbCrLf & "Renames the file to the title of the file (no longer renaming your  UPLOAD34132413241234.zip to normal names" & vbCrLf & _
"" & vbCrLf & "Stores all data into a access 2000 database (With readable table names and  cool data model)" & vbCrLf & _
"" & vbCrLf & "Note:" & vbCrLf & "Please register all files in the Bin dir" & vbCrLf & _
"This code is just a BETA version." & vbCrLf & _
"It works fine for me but I do not guarantee that it will work on your system" & vbCrLf & _
"You can email me but don't think I am unemployed. I am busy with a lot of thinks!" & vbCrLf & _
"There a a lot of ocx an dll files in this program. Some of them are part of a professional application that will be released in the future that's  the reason why you don't get all the code of the dll/ocx's" & vbCrLf & _
"" & vbCrLf & "Lot of people get bugs in previous versions  some are fixed some are there Sorry" & vbCrLf


End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Label1_Click()
Navigate Me, "Mailto:JosGroen@hotmail.com"
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Label1.ForeColor = RGB(255, 255, 255)
Me.tmrHighlight.Enabled = True

End Sub

Private Sub tmrHighlight_Timer()
Me.Label1.ForeColor = RGB(0, 0, 0)
Me.tmrHighlight.Enabled = False
End Sub
