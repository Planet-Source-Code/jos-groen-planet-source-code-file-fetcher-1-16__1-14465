VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPSC 
   Caption         =   "tblPS_PSC"
   ClientHeight    =   7095
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11610
   Icon            =   "frmPSC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7095
   ScaleWidth      =   11610
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   11610
      TabIndex        =   42
      Top             =   6000
      Width           =   11610
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   47
         Top             =   0
         Width           =   1980
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   2448
         TabIndex        =   46
         Top             =   0
         Width           =   1980
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   4837
         TabIndex        =   45
         Top             =   0
         Width           =   1980
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   7226
         TabIndex        =   44
         Top             =   0
         Width           =   1980
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   9615
         TabIndex        =   43
         Top             =   0
         Width           =   1980
      End
   End
   Begin VB.Data datPrimaryRS 
      Align           =   2  'Align Bottom
      Caption         =   "datPrimaryRS"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6330
      Width           =   11610
   End
   Begin VB.PictureBox PicStatusBar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   774
      TabIndex        =   37
      Top             =   6750
      Width           =   11610
      Begin VB.PictureBox ctrlAnimation 
         Height          =   375
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   675
         TabIndex        =   49
         Top             =   0
         Width           =   735
      End
      Begin VB.Timer tmrStatusBar1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   7200
         Top             =   105
      End
      Begin VB.Timer tmrStatusBar 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   6645
         Top             =   60
      End
      Begin VB.PictureBox ctrlProgressBar1 
         Height          =   285
         Left            =   9075
         ScaleHeight     =   225
         ScaleWidth      =   2445
         TabIndex        =   39
         Top             =   45
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.Label lblStatusbar 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   570
         TabIndex        =   40
         Top             =   75
         Width           =   75
      End
      Begin VB.Label lblStatusbar1 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   6180
         TabIndex        =   38
         Top             =   75
         Visible         =   0   'False
         Width           =   45
      End
   End
   Begin VB.Frame fraSearch 
      Height          =   540
      Left            =   4470
      TabIndex        =   33
      Top             =   -90
      Width           =   3645
      Begin VB.CommandButton cmdNormalSearchwindows 
         Height          =   300
         Left            =   3165
         Picture         =   "frmPSC.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   165
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Go"
         Height          =   300
         Left            =   2745
         TabIndex        =   36
         Top             =   165
         Width           =   375
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Left            =   990
         TabIndex        =   34
         Top             =   165
         Width           =   1740
      End
      Begin VB.Label lblSearch 
         Caption         =   "Search for:  "
         Height          =   225
         Left            =   90
         TabIndex        =   35
         Top             =   210
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdToday 
      Caption         =   "New Today"
      Height          =   345
      Left            =   105
      TabIndex        =   31
      Top             =   525
      Width           =   1365
   End
   Begin VB.Frame Frame2 
      Height          =   5025
      Left            =   9000
      TabIndex        =   28
      Top             =   975
      Width           =   2475
      Begin MSDBGrid.DBGrid DBCompatability 
         Bindings        =   "frmPSC.frx":0C40
         Height          =   2820
         Left            =   120
         OleObjectBlob   =   "frmPSC.frx":0C5F
         TabIndex        =   29
         Top             =   1560
         Width           =   2280
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmPSC.frx":163A
         Height          =   450
         Left            =   105
         OleObjectBlob   =   "frmPSC.frx":1654
         TabIndex        =   30
         Top             =   195
         Width           =   2280
      End
   End
   Begin VB.Data datCompatability 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8940
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   30
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options"
      Height          =   360
      Left            =   3015
      TabIndex        =   27
      Top             =   60
      Width           =   1410
   End
   Begin VB.CommandButton cmdAnalyse 
      Caption         =   "Fetch Mail"
      Height          =   360
      Left            =   90
      TabIndex        =   26
      Top             =   75
      Width           =   1410
   End
   Begin VB.Data datCategory 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7860
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   30
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Frame frmRecord 
      Height          =   5025
      Left            =   60
      TabIndex        =   1
      Top             =   975
      Width           =   8955
      Begin VB.Frame Frame1 
         Height          =   450
         Left            =   60
         TabIndex        =   21
         Top             =   3795
         Width           =   8850
         Begin VB.CommandButton cmdInternet 
            Height          =   285
            Left            =   8175
            Picture         =   "frmPSC.frx":2027
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Jump to internet page"
            Top             =   120
            Width           =   315
         End
         Begin VB.TextBox txtFields 
            DataField       =   "PS_HTTP"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   6
            Left            =   435
            TabIndex        =   23
            Top             =   135
            Width           =   7725
         End
         Begin VB.CommandButton cmdDownloadFile 
            Height          =   285
            Left            =   8490
            Picture         =   "frmPSC.frx":23F4
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Download file"
            Top             =   120
            Width           =   315
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "URL:"
            Height          =   195
            Index           =   6
            Left            =   45
            TabIndex        =   24
            Top             =   180
            Width           =   375
         End
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H8000000B&
         DataField       =   "PS_ID"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   411
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   200
         Width           =   510
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PS_TITLE"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   1320
         TabIndex        =   10
         Top             =   600
         Width           =   7530
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PS_DESCRIPTION"
         DataSource      =   "datPrimaryRS"
         Height          =   2685
         Index           =   2
         Left            =   1305
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1050
         Width           =   7530
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PS_DDReceived"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   3
         Left            =   1563
         TabIndex        =   8
         Top             =   200
         Width           =   1620
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PS_DDSUBMITTED"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   4
         Left            =   4335
         TabIndex        =   7
         Top             =   200
         Width           =   960
      End
      Begin VB.CheckBox chkFields 
         DataField       =   "PS_DOWNLOADED"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   5
         Left            =   6687
         TabIndex        =   6
         Top             =   200
         Width           =   255
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PS_LEVEL"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   7
         Left            =   7650
         TabIndex        =   5
         Top             =   210
         Width           =   1245
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PS_LOCALDIR"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   8
         Left            =   1440
         TabIndex        =   4
         Top             =   4650
         Width           =   7140
      End
      Begin VB.TextBox txtSERVERDIR 
         DataField       =   "PS_SERVERFILE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   4290
         Width           =   7455
      End
      Begin VB.CommandButton cmdOpenFile 
         Height          =   285
         Left            =   8580
         Picture         =   "frmPSC.frx":253E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Open the file with his default viewer"
         Top             =   4650
         Width           =   315
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   8895
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   225
         Width           =   210
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "TITLE:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   19
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1050
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "SENT:"
         Height          =   195
         Index           =   3
         Left            =   1002
         TabIndex        =   17
         Top             =   225
         Width           =   480
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "SUBMITTED:"
         Height          =   195
         Index           =   4
         Left            =   3264
         TabIndex        =   16
         Top             =   225
         Width           =   990
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "DOWNLOADED:"
         Height          =   195
         Index           =   5
         Left            =   5376
         TabIndex        =   15
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "LEVEL:"
         Height          =   195
         Index           =   7
         Left            =   7023
         TabIndex        =   14
         Top             =   225
         Width           =   540
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "LOCALFILENAME:"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   13
         Top             =   4680
         Width           =   1350
      End
      Begin VB.Label lblLServerDIr 
         AutoSize        =   -1  'True
         Caption         =   "FILE ON SERVER"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   4335
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdDownloadAllFiles 
      Caption         =   "Download All"
      Height          =   360
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "Download all files if downloaded is false"
      Top             =   75
      Width           =   1410
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5205
      Top             =   3300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPSC.frx":2674
            Key             =   "Word Underline"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPSC.frx":2786
            Key             =   "Undo"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblVesion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "V1.2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   10920
      TabIndex        =   41
      Top             =   750
      Width           =   435
   End
   Begin VB.Label lblInfo 
      Caption         =   "JosGroen@Hotmail.com"
      Height          =   240
      Left            =   8370
      TabIndex        =   32
      Top             =   750
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   10095
      Picture         =   "frmPSC.frx":2898
      Top             =   30
      Width           =   1410
   End
End
Attribute VB_Name = "frmPSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Public bFirstTime As Boolean
Dim m_lRecCount As Long


Private Sub cmdDownloadAllFiles_Click()

    DisplayInfoCls
    DisplayInfo "Download all files"
    DoEvents
    With datPrimaryRS.Recordset
        .MoveFirst
        Do While Not .EOF
            If chkFields(5).Value = 0 Then
                Call DisplayInfoProgressbar1(Me.datPrimaryRS.Recordset.AbsolutePosition, m_lRecCount)
                cmdDownloadFile_Click
            End If
            DoEvents
            .MoveNext
            
            Call DisplayInfoProgressbar1(Me.datPrimaryRS.Recordset.AbsolutePosition, m_lRecCount)

        Loop
        .MoveFirst
        DisplayInfoProgressbar1Hide
    End With
 
End Sub

Private Sub cmdDownloadFile_Click()

  Dim strFilename As String
  Dim strPath As String
  Dim strNewpath As String

    Call DisplayIconInit(AnimationType.Seaching, 100)
    Call DisplayInfoLbl("Searching for file")
    strFilename = Me.datPrimaryRS.Recordset("PS_TITLE")                '.txtFields(1)
    strFilename = Replace$(strFilename, """", "_")
    strFilename = Replace$(strFilename, "'", "_")
    strFilename = Replace$(strFilename, "/", "_")
    strFilename = Replace$(strFilename, "\", "_")
    strFilename = Replace$(strFilename, 0, "_", , , vbBinaryCompare)
    strFilename = Replace$(strFilename, " ", "_")
    strFilename = Replace$(strFilename, ":", "_")
    strFilename = Replace$(strFilename, ".", "_")
    strFilename = Replace$(strFilename, "*", "_")
    strFilename = Replace$(strFilename, "<", "_")
    strFilename = Replace$(strFilename, ">", "_")
    strFilename = Replace$(strFilename, "?", "_")
AskAgain:
    strPath = Reg.PM_GetRegValue("Options", "DefaultDownloadDir")
    Set File = New J2P_FILE.clsFile
    
    If File.PathExists(strPath) = J2P_FILE.EnError.Path_Not_Exists Then
        strNewpath = InputBox("The Default Download dir is invalid" & vbLf & _
                     "Please give a valid path", , strPath)
        If strNewpath = "" Then
            Exit Sub
          Else
            If File.PathExists(strNewpath) = J2P_FILE.EnError.Path_Not_Exists Then
                GoTo AskAgain
              Else
                Call Reg.PM_SetRegValue("Options", "DefaultDownloadDir", strNewpath)
                GoTo AskAgain
            End If
        
        End If
    
    End If
    
    strPath = Left$(strPath, Len(strPath) - 2)

    strPath = strPath & "\" & Trim$(strFilename)
    
 '   Load frmDownload
    Me.datPrimaryRS.Recordset.Edit
    Me.txtSERVERDIR = frmDownload.GetDownloadUrl(Me.txtFields(6))

    If Me.txtSERVERDIR <> "<NotFound>" Then
 
        Me.txtFields(8) = strPath & ".zip"
        Call frmDownload.DownloadFile(Me.txtSERVERDIR, Me.txtFields(8))
        
      Else
        Me.txtFields(8) = strPath & ".htm"
        Call frmDownload.DownloadFile(Me.txtFields(6), Me.txtFields(8))
        'DisplayInfo "Downloaded as htm"
    End If
'    Unload frmDownload
    chkFields(5).Value = 1
    
    Call DisplayIconHide
    cmdUpdate_Click

End Sub

Private Sub cmdInternet_Click()

    Navigate Me, Me.datPrimaryRS.Recordset("PS_HTTP")

End Sub

Private Sub cmdNormalSearchwindows_Click()
If frmSearch.WindowState = vbMinimized Then
    frmSearch.WindowState = vbNormal
    Me.cmdNormalSearchwindows.Visible = False
End If


End Sub

Private Sub cmdOpenFile_Click()

    Navigate Me, Me.txtFields(8)

End Sub

Private Sub cmdOptions_Click()

    frmOptions.Show , Me

End Sub

Private Sub cmdSearch_Click()

    Call DisplayInfoLbl("Searching titles and descriptions....")
    Call DisplayIconInit(AnimationType.Search)
    
    If Len(Trim$(Me.txtSearch.Text)) > 0 Then
        frmSearch.p_strToSearchFor = Trim$(Me.txtSearch.Text)

        If frmSearch.Visible = False Then
            Load frmSearch
            frmSearch.Show
          Else
            
            'Call frmSearch.cmdSearch_Click
            Unload frmSearch
            Load frmSearch
            frmSearch.Show
            
        End If
        
      Else
        Call DisplayIconHide
        MsgBox "Empty string not allowed in search box"
        
    End If

End Sub

Private Sub cmdToday_Click()

    frmNewToday.Show

End Sub


Private Sub datPrimaryRS_Reposition()

  Dim cSql As String

    'This will display the current record position for this recordset
    datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
    Me.frmRecord.Caption = CStr(datPrimaryRS.Recordset.AbsolutePosition) & "/" & CStr(m_lRecCount)

    'Catogory
    If (Not datPrimaryRS.Recordset.EOF) And (Not datPrimaryRS.Recordset.BOF) Then
        If Not IsNull(datPrimaryRS.Recordset("PS_ID")) Then
            cSql = "SELECT CA_CategoryName as Catogory FROM tblCA_CategoryNames INNER JOIN tblCC_CategoryWith ON tblCA_CategoryNames.CA_ID = tblCC_CategoryWith.CC_CA_ID " & _
                   " Where (((tblCC_CategoryWith.CC_PS_ID) = " & CStr(datPrimaryRS.Recordset("PS_ID")) & ")) " & _
                   " ORDER BY tblCC_CategoryWith.CC_PS_ID; "
            Me.datCategory.RecordSource = cSql
          Else
            Me.datCategory.RecordSource = ""
        End If
        Me.datCategory.Refresh
        'Compatability
        If Not IsNull(datPrimaryRS.Recordset("PS_ID")) Then
        
            cSql = "SELECT tblCN_CompatabilityNames.CN_CompatibilityName as Compatability " & _
                   "FROM tblCN_CompatabilityNames INNER JOIN tblCW_CompatabilityWith ON tblCN_CompatabilityNames.CN_ID = tblCW_CompatabilityWith.CW_CN_ID " & _
                   "Where (((tblCW_CompatabilityWith.CW_PS_ID) = " & CStr(datPrimaryRS.Recordset("PS_ID")) & ")) " & _
                   "ORDER BY tblCN_CompatabilityNames.CN_CompatibilityName DESC; "
        
            Me.datCompatability.RecordSource = cSql
          Else
            Me.datCompatability.RecordSource = ""
        End If
        
        Me.datCompatability.Refresh
        Me.DBGrid1.Height = (Me.DBGrid1.ApproxCount + 2) * Me.DBGrid1.RowHeight
        Me.DBCompatability.Height = (Me.DBCompatability.ApproxCount + 2) * Me.DBCompatability.RowHeight
        
        Me.DBCompatability.Top = Me.DBGrid1.Height + Me.DBGrid1.Top
    End If

End Sub

Private Sub DBGrid1_Click()

    Me.DBGrid1.Height = (Me.DBGrid1.ApproxCount + 2) * Me.DBGrid1.RowHeight

End Sub

Private Sub Form_Load()
    
    Me.ctrlProgressBar1.LoadMeter (RGB(64, 64, 128))
    '    Me.datPrimaryRS.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & gStrDatabaseFilename & ";"
    Me.datPrimaryRS.DatabaseName = gStrDatabaseFilename
    Me.datPrimaryRS.RecordSource = "SELECT PS_ID, PS_LEVEL, PS_DDSUBMITTED, PS_DDReceived, PS_HTTP, PS_LOCALDIR, PS_CODEOFTHEDAY, PS_DOWNLOADED, PS_TITLE, PS_MSGID, PS_CATOGORY, PS_Compatibility, PS_DESCRIPTION, PS_SERVERFILE FROM tblPS_PSC Order by PS_DDReceived DESC"
    Me.datCategory.DatabaseName = gStrDatabaseFilename
    Me.datCompatability.DatabaseName = gStrDatabaseFilename
    Me.datCategory.Refresh
    Me.datCompatability.Refresh
    Me.datPrimaryRS.Refresh
    Me.Caption = g_ApplTitle & " - Main"
    Me.lblVesion.Caption = "V" & App.Major & "." & App.Minor
 '   If Not datPrimaryRS.Recordset.EOF Then
 '       datPrimaryRS.Recordset.MoveLast
'    End If
    m_lRecCount = datPrimaryRS.Recordset.AbsolutePosition
    
    If Not datPrimaryRS.Recordset.BOF Then
        datPrimaryRS.Recordset.MoveFirst
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload frmMain
    Unload frmDialog
    Unload frmSearch
    Screen.MousePointer = vbDefault

End Sub

'Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

'This is where you would put error handling code
'If you want to ignore errors, comment out the next line
'If you want to trap them, add code here to handle them

'End Sub



Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As adodb.EventReasonEnum, ByVal cRecords As Long, adStatus As adodb.EventStatusEnum, ByVal pRecordset As adodb.Recordset)

  'This is where you put validation code
  'This event gets called when the following actions occur
  
  Dim bCancel As Boolean

    Select Case adReason
      Case adRsnAddNew
      Case adRsnClose
      Case adRsnDelete
      Case adRsnFirstChange
      Case adRsnMove
      Case adRsnRequery
      Case adRsnResynch
      Case adRsnUndoAddNew
      Case adRsnUndoDelete
      Case adRsnUndoUpdate
      Case adRsnUpdate
    End Select

    If bCancel Then
        adStatus = adStatusCancel
    End If

End Sub

Private Sub cmdAdd_Click()

    On Error GoTo AddErr
    datPrimaryRS.Recordset.AddNew

Exit Sub

AddErr:
    MsgBox Err.Description

End Sub

Private Sub cmdDelete_Click()

    On Error GoTo DeleteErr
    With datPrimaryRS.Recordset
        .Delete
        .MoveNext
        If .EOF Then
            .MoveLast
        End If
    End With

Exit Sub

DeleteErr:
    MsgBox Err.Description

End Sub

Private Sub cmdRefresh_Click()

  'This is only needed for multi user apps

    On Error GoTo RefreshErr
    datPrimaryRS.Refresh

Exit Sub

RefreshErr:
    MsgBox Err.Description

End Sub

Private Sub cmdUpdate_Click()

    On Error GoTo UpdateErr

    datPrimaryRS.Recordset.Update 'Batch adAffectAll

Exit Sub

UpdateErr:
    If Err.Number = 3020 Then 'update without addnew or edit
        datPrimaryRS.Recordset.Edit
        Resume
      Else
        MsgBox Err.Description & Err.Number
    End If

End Sub

Private Sub cmdClose_Click()
    
    Unload Me

End Sub

Private Sub cmdAnalyse_Click()

    Call DisplayIconInit(AnimationType.Email)
    Call DisplayInfoLbl("Analyse mailbox")

    DisplayInfo "Fetch Mail"
    Call Uitkijk_Logon(frmMain.MAPISession, frmMain.MAPIMessages)
    Call Uitkijk_Analyse
    DisplayIconHide
    Me.datPrimaryRS.Recordset.Requery

End Sub

Private Sub Image2_Click()

End Sub

Private Sub tmrStatusBar_Timer()

    m_lngTmrCleanUptext = m_lngTmrCleanUptext + 1
    If m_lngTmrCleanUptext > 50 And frmPSC.ctrlAnimation.Visible = False Then
        tmrStatusBar.Enabled = False
        Me.lblStatusbar.Caption = "Done."
   
    End If

End Sub

Private Sub tmrStatusBar1_Timer()

    m_lngTmrCleanUptext1 = m_lngTmrCleanUptext1 + 1
    If m_lngTmrCleanUptext1 = 50 Then
        tmrStatusBar.Enabled = False
        Me.lblStatusbar1.Visible = False
        Me.lblStatusbar1.Caption = ""
    End If

End Sub


