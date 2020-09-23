VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4920
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6150
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   4080
      Left            =   105
      TabIndex        =   9
      Top             =   180
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7197
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmOptions.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkFeedback"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameDatabase"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraScandate"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CommonDialog1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.Frame Frame2 
         Caption         =   "Log file"
         Height          =   720
         Left            =   165
         TabIndex        =   20
         Top             =   2790
         Width           =   5565
         Begin VB.TextBox txtLogfile 
            Height          =   345
            Left            =   120
            TabIndex        =   22
            Top             =   270
            Width           =   4995
         End
         Begin VB.CommandButton cmdGetlogfile 
            Height          =   345
            Left            =   5145
            MaskColor       =   &H00C0C000&
            Picture         =   "frmOptions.frx":0028
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   270
            Width           =   360
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   75
         Top             =   3495
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame1 
         Caption         =   "Database"
         Height          =   720
         Left            =   165
         TabIndex        =   16
         Top             =   435
         Width           =   5565
         Begin VB.TextBox txtDatabasepath 
            Height          =   345
            Left            =   120
            TabIndex        =   18
            Top             =   270
            Width           =   4995
         End
         Begin VB.CommandButton cmdGivePointToDatabase 
            Height          =   345
            Left            =   5160
            MaskColor       =   &H00C0C000&
            Picture         =   "frmOptions.frx":017A
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   270
            Width           =   360
         End
      End
      Begin VB.Frame fraScandate 
         Caption         =   "Scan only messages newer than"
         Height          =   720
         Left            =   165
         TabIndex        =   14
         Top             =   2005
         Width           =   5565
         Begin VB.CheckBox ChkAutoUpdateDate 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto update:"
            Height          =   195
            Left            =   4065
            TabIndex        =   19
            Top             =   330
            Width           =   1350
         End
         Begin MSComCtl2.DTPicker txtScandate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "d.MMMM yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   3
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   15
            Top             =   270
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   609
            _Version        =   393216
            Format          =   24576000
            CurrentDate     =   36849
         End
      End
      Begin VB.Frame FrameDatabase 
         Caption         =   "Default download path"
         Height          =   720
         Left            =   165
         TabIndex        =   11
         Top             =   1220
         Width           =   5565
         Begin VB.CommandButton cmdDownloadfolder 
            Height          =   345
            Left            =   5145
            MaskColor       =   &H00C0C000&
            Picture         =   "frmOptions.frx":02CC
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   270
            Width           =   360
         End
         Begin VB.TextBox txtDownloadDir 
            Height          =   345
            Left            =   120
            TabIndex        =   12
            Top             =   270
            Width           =   4995
         End
      End
      Begin VB.CheckBox chkFeedback 
         Alignment       =   1  'Right Justify
         Caption         =   "Feedback"
         Height          =   315
         Left            =   3900
         TabIndex        =   10
         Top             =   2805
         Visible         =   0   'False
         Width           =   1665
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2490
      TabIndex        =   0
      Top             =   4455
      Width           =   1095
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkAutoUpdateDate_Click()
Me.cmdApply.Enabled = True
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdGetlogfile_Click()

  Dim strResFolder As String

    strResFolder = BrowseForFolder(hwnd, "Select Logfile Folder" & vbLf & "Where the logfile will be created")
    Me.txtLogfile.Text = strResFolder
    Me.cmdApply.Enabled = True

End Sub

Private Sub Form_Load()
    Dim intResult As Integer
    Dim VarReturnValue As Variant
    Me.txtScandate.Value = g_datScan
    Me.Caption = g_ApplTitle & " - " & "Settings"
    Me.chkFeedback.Value = IIf(g_bFeedback = True, 1, 0)
    intResult = Reg.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "DefaultDownloadDir", VarReturnValue)
    Me.txtDownloadDir.Text = VarReturnValue
    intResult = Reg.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "Feedback", VarReturnValue)
    Me.chkFeedback.Value = VarReturnValue
    intResult = Reg.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "DatabaseLocation", VarReturnValue)
    Me.txtDatabasepath.Text = VarReturnValue
    VarReturnValue = 0
    intResult = Reg.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "Scandate", VarReturnValue)
    Me.txtScandate.Value = VarReturnValue
    VarReturnValue = 0
    intResult = Reg.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "AutoUpdateDate", VarReturnValue)
    Me.ChkAutoUpdateDate.Value = VarReturnValue

    intResult = Reg.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "Logfile", VarReturnValue)
   Debug.Print VarReturnValue
   Me.txtLogfile.Text = IIf(VarReturnValue = 0, App.Path & "\" & "PSCLOGFILE.TXT", CStr(VarReturnValue))

End Sub

Private Sub cmdApply_Click()
    Dim intResult As Integer
    intResult = Reg.CreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options")
    
    intResult = Reg.SetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "DefaultDownloadDir", Me.txtDownloadDir.Text)
    intResult = Reg.SetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "Feedback", Me.chkFeedback.Value)
    intResult = Reg.SetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "DatabaseLocation", Me.txtDatabasepath.Text)
    
    intResult = Reg.SetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "Scandate", CStr(Me.txtScandate.Value))
    
    intResult = Reg.SetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "AutoUpdateDate", Me.ChkAutoUpdateDate.Value)
    
    intResult = Reg.SetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "Logfile", Me.txtLogfile.Text)
    g_Logfile = Me.txtLogfile.Text
     
    gStrDefaultDownloadDir = Me.txtDownloadDir.Text
    g_datScan = Me.txtScandate.Value
    Me.cmdApply.Enabled = False
End Sub












Private Sub chkFeedback_Click()
g_bFeedback = IIf(Me.chkFeedback.Value = 1, True, False)
End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdGivePointToDatabase_Click()
Dim strReturnedFilename As String
With Me.CommonDialog1
.DialogTitle = g_ApplTitle & "- Point to database"
.Filter = "Microsoft Access Database (*.mdb)|*.mdb|"
.ShowOpen
strReturnedFilename = .FileName
End With

'did user select a file ?
If IsNull(strReturnedFilename) = False Then
    Me.txtDatabasepath.Text = strReturnedFilename
    Me.cmdApply.Enabled = True
End If

End Sub

Private Sub cmdOK_Click()
'    MsgBox "Place code here to set options and close dialog!"
    Unload Me
End Sub



Private Sub cmdDownloadfolder_Click()

  Dim strResFolder As String

    strResFolder = BrowseForFolder(hwnd, "Please select a folder.")
    Me.txtDownloadDir.Text = strResFolder
    Me.cmdApply.Enabled = True

End Sub





Private Sub txtDownloadDir_Change()
Me.cmdApply.Enabled = True
End Sub

Private Sub txtScandate1_Change()
Me.cmdApply.Enabled = True
End Sub
