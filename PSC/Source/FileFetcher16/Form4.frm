VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D6BB5C89-849D-4016-BD58-6357C24DC813}#2.1#0"; "PSC_Toolbar.ocx"
Object = "{121E2F4F-F2DF-410E-9CD4-956E95AC2412}#1.0#0"; "PSC_PROGRESSBAR.OCX"
Object = "{8A66605F-330F-4870-B2F6-44E93996096E}#2.0#0"; "J2_Treeview.ocx"
Object = "{AB5D4558-C297-440D-91CB-5DFA4501926E}#1.0#0"; "PSC_GLOBE.OCX"
Object = "{853443F2-262B-43D7-A139-F20E3795008B}#2.1#0"; "PSC_Listview.ocx"
Object = "{787BAFE2-B2B1-4380-86F3-D4576AACE74B}#1.1#0"; "PSC_Details.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11895
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin PSC_Details.ctlDetails ctlDetails1 
      Height          =   6270
      Left            =   4080
      TabIndex        =   20
      Top             =   1080
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   11060
   End
   Begin Globe.ctlGlobe ctlGlobe1 
      Height          =   555
      Left            =   11190
      TabIndex        =   17
      Top             =   30
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PSC_Toolbar.ctlToolbar Toolbar1 
      Height          =   690
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   1217
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picStatusbar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   11895
      TabIndex        =   3
      Top             =   7800
      Width           =   11895
      Begin PSC_Progressbar.ctrlProgressBar ctrlProgressBar1 
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   225
         Width           =   12105
         _ExtentX        =   21352
         _ExtentY        =   450
      End
      Begin VB.Label lblDisplayinfo1 
         Caption         =   "Done..."
         Height          =   210
         Left            =   60
         TabIndex        =   12
         Top             =   0
         Width           =   12030
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":08CA
            Key             =   "email"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":0F56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3950
      _ExtentX        =   6959
      _ExtentY        =   11668
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Category"
      TabPicture(0)   =   "Form4.frx":10B2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtSearchstr"
      Tab(0).Control(1)=   "CmdSearch"
      Tab(0).Control(2)=   "ctlTreeview1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Compatibility"
      TabPicture(1)   =   "Form4.frx":10CE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ctlTreeview2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Last download"
      TabPicture(2)   =   "Form4.frx":10EA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ctlListview1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Search"
      TabPicture(3)   =   "Form4.frx":1106
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Picture1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin PSC_Listview.ctlListview ctlListview1 
         Height          =   6105
         Left            =   -74925
         TabIndex        =   18
         Top             =   450
         Width           =   3800
         _ExtentX        =   6694
         _ExtentY        =   10769
      End
      Begin J2_Treeview.ctlTreeview ctlTreeview2 
         Height          =   6060
         Left            =   -74925
         TabIndex        =   15
         Top             =   480
         Width           =   3800
         _ExtentX        =   6694
         _ExtentY        =   10689
         Appearance      =   1
      End
      Begin J2_Treeview.ctlTreeview ctlTreeview1 
         Height          =   5745
         Left            =   -74925
         TabIndex        =   14
         Top             =   780
         Width           =   3800
         _ExtentX        =   6694
         _ExtentY        =   10134
         Appearance      =   1
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6075
         Left            =   75
         ScaleHeight     =   6045
         ScaleWidth      =   3765
         TabIndex        =   8
         Top             =   480
         Width           =   3800
         Begin PSC_Listview.ctlListview ctlListview2 
            Height          =   4605
            Left            =   45
            TabIndex        =   16
            Top             =   1335
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   8123
            BackStyle       =   0
            Appearance      =   0
         End
         Begin VB.CommandButton cmdSearchNow 
            Caption         =   "Search Now"
            Height          =   315
            Left            =   45
            TabIndex        =   5
            Top             =   795
            Width           =   1260
         End
         Begin VB.TextBox txtSearch 
            Height          =   315
            Left            =   50
            TabIndex        =   10
            Top             =   465
            Width           =   3360
         End
         Begin VB.Label lblFound 
            BackStyle       =   0  'Transparent
            Height          =   210
            Left            =   675
            TabIndex        =   19
            Top             =   1125
            Width           =   2760
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Results:"
            Height          =   195
            Left            =   45
            TabIndex        =   6
            Top             =   1125
            Width           =   600
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Search for program named:"
            Height          =   255
            Left            =   50
            TabIndex        =   11
            Top             =   255
            Width           =   3195
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Search for programs"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   315
            TabIndex        =   9
            Top             =   60
            Width           =   3135
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   50
            Picture         =   "Form4.frx":1122
            Top             =   15
            Width           =   210
         End
      End
      Begin VB.CommandButton CmdSearch 
         Height          =   300
         Left            =   -71415
         Picture         =   "Form4.frx":1424
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   450
         Width           =   300
      End
      Begin VB.TextBox txtSearchstr 
         Height          =   300
         Left            =   -74925
         TabIndex        =   2
         Text            =   "<Enter Search Text>"
         Top             =   465
         Width           =   3510
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10470
      Top             =   45
      Width           =   645
   End
   Begin VB.Label lblDisplayInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4065
      TabIndex        =   13
      Top             =   735
      Width           =   7785
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   7650
      X2              =   11790
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   7650
      X2              =   11745
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Settings 
         Caption         =   "&Settings"
      End
   End
   Begin VB.Menu About 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnbuildCategory As Boolean
Dim blnbuildCompatibility As Boolean
Dim blnbuildToday As Boolean

Dim blnExpandedTree1 As Boolean     ' is category uitgeklapt ?
Dim blnExpandedTree2 As Boolean     ' is compatibiliy uitgeklapt ?

Dim blnExpandAll As Boolean

Private WithEvents m_TreeEvents As J2_Treeview.ctlTreeview
Attribute m_TreeEvents.VB_VarHelpID = -1
Private WithEvents m_TreeEvents1 As J2_Treeview.ctlTreeview
Attribute m_TreeEvents1.VB_VarHelpID = -1
Private WithEvents m_TreeEvents2 As J2_Treeview.ctlTreeview
Attribute m_TreeEvents2.VB_VarHelpID = -1
Private WithEvents m_Details As psc_details.ctlDetails
Attribute m_Details.VB_VarHelpID = -1



Private Sub buildCategory()

  Dim intResult As Integer

    With Me.ctlTreeview1
        
        '- Verbinding naar de database instellen.
        .Connecttype = "Access2000;"
        .DatabaseName = CStr(gStrDatabaseFilename)
        '- Aangeven welke velden en Query voor de Nodes
        .txtRoot = "ROOT"
        .SqlQuery = "SELECT tblCA_CategoryNames.CA_ID AS NODE, tblCA_CategoryNames.CA_CategoryName AS TEKST, 1 AS PARENT FROM tblCA_CategoryNames ORDER BY tblCA_CategoryNames.CA_CategoryName;"
        .fldNode = "NODE"
        .fldParent = "PARENT"
        .fldNodetext = "TEKST"
        '- Aangeven welke velden en Query voor de Leafs.
        .SqlQueryLeaf = "SELECT tblCC_CategoryWith.CC_PS_ID AS NODE, tblCA_CategoryNames.CA_ID AS PARENT, tblPS_PSC.PS_TITLE AS TEKST FROM tblCA_CategoryNames INNER JOIN (tblCC_CategoryWith INNER JOIN tblPS_PSC ON tblCC_CategoryWith.CC_PS_ID = tblPS_PSC.PS_ID) ON tblCA_CategoryNames.CA_ID = tblCC_CategoryWith.CC_CA_ID;"
        .fldNodeleaf = "NODE"
        .fldParentleaf = "PARENT"
        .fldNodetextleaf = "TEKST"
        .IconSet = J2_Treeview.enIcons.VbFiles
        '- Aanwezige nodes verwijderen.
        .Clear
        '- Daadwerkelijk opbouwen van de Tree
        intResult = .Populate
        .SortAll
        .ExpandLevelOne
        
    End With

End Sub

Private Sub buildCompatibility()
  Dim intResult As Integer

    With Me.ctlTreeview2
        
        '- Verbinding naar de database instellen.
        .Connecttype = "Access2000;"
        .DatabaseName = CStr(gStrDatabaseFilename)
        '- Aangeven welke velden en Query voor de Nodes
        .txtRoot = "ROOT"
        .SqlQuery = "SELECT tblCN_CompatabilityNames.CN_ID AS NODE, tblCN_CompatabilityNames.CN_CompatibilityName AS TEKST, 1 AS PARENT FROM tblCN_CompatabilityNames ORDER BY tblCN_CompatabilityNames.CN_CompatibilityName;"
        .fldNode = "NODE"
        .fldParent = "PARENT"
        .fldNodetext = "TEKST"
        
        '- Aangeven welke velden en Query voor de Leafs.
       
        .SqlQueryLeaf = "SELECT tblCW_CompatabilityWith.CW_PS_ID AS NODE, tblCN_CompatabilityNames.CN_ID AS PARENT, tblPS_PSC.PS_TITLE AS TEKST FROM tblCN_CompatabilityNames INNER JOIN (tblPS_PSC INNER JOIN tblCW_CompatabilityWith ON tblPS_PSC.PS_ID = tblCW_CompatabilityWith.CW_PS_ID) ON tblCN_CompatabilityNames.CN_ID = tblCW_CompatabilityWith.CW_CN_ID ORDER BY tblPS_PSC.PS_TITLE;"
        .fldNodeleaf = "NODE"
        .fldParentleaf = "PARENT"
        .fldNodetextleaf = "TEKST"
        .IconSet = J2_Treeview.enIcons.VbFiles
        '- Aanwezige nodes verwijderen.
        .Clear
        intResult = .Populate
        .SortAll
        .ExpandLevelOne
    
    End With

End Sub

Private Sub About_Click()
frmSplash.Show
End Sub

Private Sub cmdSearch_Click()
    GoSearch
End Sub

Private Sub DownloadAll()
On Error Resume Next
Me.WindowState = vbMinimized
  Dim iResult As Integer
  Dim rstCom As adodb.Recordset
  Dim strSql As String
  Dim lngRecTeller As Long
  Dim lngrecCount As Long
  Dim sTime As Date
    strSql = "SELECT PS_ID,PS_HTTP,PS_Title FROM tblPS_PSC WHERE PS_DOWNLOADED = 'N' ORDER BY PS_DDReceived DESC "
   
    iResult = Ado.J2_Connect(gStrDatabaseFilename, Access2000)
    If iResult = J2_ADO.EnError.No_Errors Then
        iResult = Ado.J2_Recordset(strSql, adOpenKeyset, adLockOptimistic, adCmdText, rstCom)
        If iResult = J2_ADO.EnError.No_Errors Then
            If Not (rstCom.EOF And rstCom.BOF) Then
              rstCom.MoveLast
              lngrecCount = rstCom.RecordCount
              rstCom.MoveFirst
              
              Do While rstCom.EOF = False
                sTime = Now()
                lngRecTeller = lngRecTeller + 1
                DisplayInfoProgressbar1 lngRecTeller, lngrecCount
                Me.Caption = CStr(lngRecTeller) & "/" & CStr(lngrecCount)
                Call File.HORAEST_AddLineToFile(g_Logfile, "start: " & CStr(sTime) & " " & CStr(lngRecTeller) & " " & rstCom("PS_Title").Value)
                Call frmDownload.GetDownloadUrl(rstCom("PS_ID").Value, rstCom("PS_HTTP").Value, rstCom("PS_Title").Value)
                Call File.HORAEST_AddLineToFile(g_Logfile, " Done: " & Now())
                rstCom.MoveNext
              Loop
            End If
        Else
    End If
    End If
    
Me.WindowState = vbNormal
    On Error Resume Next
      rstCom.Close
      Set rstCom = Nothing
        Ado.J2_Disconnect
    On Error GoTo 0

End Sub


Private Sub cmdSearchNow_Click()

  Dim sTemp As String
  Dim iAantalArgs As Integer
  Dim ibrowsecompa As Integer
  Dim lCnId As Long           'compatibity names id
  Dim intCountAnds As Integer
  Dim strQuery As String

    sTemp = Replace$(Me.txtSearch.Text, " ", ";")
    sTemp = Replace$(sTemp, "  ", " ")
    
    iAantalArgs = STRING_CountArgs(sTemp)
   
    'Opslaan in een array
    ReDim g_vSearch(iAantalArgs)
    g_vSearchCount = iAantalArgs
     
    'cOUNT ARGS
    For ibrowsecompa = 1 To iAantalArgs
        
        g_vSearch(ibrowsecompa) = Format(Trim$(STRING_GeefArg(ibrowsecompa, sTemp)), ">")
    
    Next ibrowsecompa

    ''Analyse string

    For ibrowsecompa = 1 To iAantalArgs
        If g_vSearch(ibrowsecompa) = "AND" Then
            intCountAnds = intCountAnds + 1
        End If
    Next ibrowsecompa

    strQuery = "SELECT tblPS_PSC.PS_ID, tblPS_PSC.PS_TITLE, tblPS_PSC.PS_DESCRIPTION From tblPS_PSC "
           
    If intCountAnds = 0 Then
        ''1 argument
        strQuery = strQuery & "WHERE (((tblPS_PSC.PS_TITLE) Like '%" & g_vSearch(1) & "%')) OR (((tblPS_PSC.PS_DESCRIPTION) Like '%" & g_vSearch(1) & "%'))"
    Else
        strQuery = strQuery & "WHERE ((tblPS_PSC.PS_TITLE) Like '%"
           
        For ibrowsecompa = 1 To iAantalArgs
            If g_vSearch(ibrowsecompa) = "AND" Then
                intCountAnds = intCountAnds + 1
            Else
                strQuery = strQuery & g_vSearch(ibrowsecompa) & "%') OR"
            End If
        Next ibrowsecompa
    End If
    
    'Order the results
    strQuery = strQuery + " Order by PS_TITLE"
    'Populate the list
    Search strQuery

End Sub



Private Sub ctlListview1_Click(ByVal Item As MSComctlLib.ListItem)
    If Left$(Item.Key, 1) = "I" Then
        ctlDetails1.PsId = Mid$(Item.Key, 2, Len(Item.Key) - 1)
    End If
End Sub

Private Sub ctlListview2_Click(ByVal Item As MSComctlLib.IListItem)
    If Left$(Item.Key, 1) = "I" Then
        ctlDetails1.PsId = Mid$(Item.Key, 2, Len(Item.Key) - 1)
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'The internet control really sucks and won't close when we ask. shit!!
    frmDownload.Inet.Cancel
    frmDownload.Inet.Cancel
    Unload frmDownload
End Sub

Private Sub Toolbar1_ButtonPress(strButton As String)
Select Case strButton

Case "btnRefresh"
    Me.Toolbar1.ButtonEnable "btnRefresh", False
    Me.ctlGlobe1.Start
    DoEvents
    buildToday
    buildCategory
    buildCompatibility
    blnbuildCategory = True
    blnbuildCompatibility = True
    blnbuildToday = True
    Me.ctlGlobe1.Off
    Me.Toolbar1.ButtonEnable "btnRefresh", True

Case "btnMail"
     Me.Toolbar1.ButtonEnable "btnMail", False
     Me.ctlGlobe1.Start
     Reconnnect
     Call Uitkijk_Logon(frmMain.MAPISession, frmMain.MAPIMessages)
     Call Uitkijk_Analyse
     Me.ctlGlobe1.Off
     Ado.J2_Disconnect
     Me.Toolbar1.ButtonEnable "btnMail", True

Case "btnPaperclip"
    Me.Toolbar1.ButtonEnable "btnPaperclip", False
    Me.ctlGlobe1.Start
    DownloadAll
    Me.ctlGlobe1.Off
    Me.Toolbar1.ButtonEnable "btnPaperclip", True

Case "btnQuestion"
    frmSplash.Show

Case "btnTv1"
    frmSettings.Show vbModal, Me

Case Else
    
End Select

End Sub

Private Sub Form_Initialize()
    buildCategory
    blnbuildCategory = True
    Me.ctlGlobe1.Off
End Sub

Private Sub Form_Load()
    Me.ctlGlobe1.Start
    
    Me.SSTab1.Tab = 0
    
    Me.ctrlProgressBar1.LoadMeter (RGB(0, 0, 128))
    Me.Caption = g_ApplTitle
   
    Set m_TreeEvents = ctlTreeview1         'Get events from Treeview
    Set m_TreeEvents1 = ctlTreeview2         'Get events from Treeview
    Set m_Details = ctlDetails1         'Get events from Treeview
   
    ctlDetails1.Database = gStrDatabaseFilename
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
    Me.SSTab1.Top = Me.Toolbar1.Top + Me.Toolbar1.Height
    Me.SSTab1.Height = Me.ScaleHeight - Me.SSTab1.Top - Me.picStatusbar.Height
    Me.ctlTreeview1.Height = Me.SSTab1.Height - Me.ctlTreeview1.Top - 100
    Me.ctlListview1.Height = Me.SSTab1.Height - Me.ctlListview1.Top - 100
    Me.ctlTreeview2.Height = Me.SSTab1.Height - Me.ctlTreeview2.Top - 100
    Me.CmdSearch.Left = Me.txtSearchstr.Left + Me.txtSearchstr.Width
On Error GoTo 0

End Sub

Private Sub m_Details_ButtonClicked(Button As psc_details.enButtonnr)
    
    Select Case Button
     Case psc_details.enButtonnr.Hyperlink
        Navigate Me, Me.ctlDetails1.AuthorHref
     Case psc_details.enButtonnr.GotoPage
        Navigate Me, Me.ctlDetails1.URL
     
     Case psc_details.enButtonnr.Openfile
        Navigate Me, Me.ctlDetails1.File
     
     Case psc_details.enButtonnr.DownloadFile
        Me.ctlGlobe1.Start
        
        Dim strGetUrl As String
        
        Call frmDownload.GetDownloadUrl(Me.ctlDetails1.PsId, Me.ctlDetails1.URL, Me.ctlDetails1.Title)
        
        On Error Resume Next
        ctlDetails1.Refresh
        Me.ctlGlobe1.Off
    
    End Select

End Sub

Private Sub m_TreeEvents_NodeClick(strNodeKey As String)
    If Left$(strNodeKey, 1) = "O" Then
        ctlDetails1.PsId = Mid$(strNodeKey, 2, InStr(1, strNodeKey, "/") - 2)
    End If
End Sub

Private Sub m_TreeEvents1_NodeClick(strNodeKey As String)
    If Left$(strNodeKey, 1) = "O" Then
        ctlDetails1.PsId = Mid$(strNodeKey, 2, InStr(1, strNodeKey, "/") - 2)
    End If
End Sub





Private Sub m_TreeEvents2_NodeClick(strNodeKey As String)
    If Left$(strNodeKey, 1) = "O" Then
        ctlDetails1.PsId = Mid$(strNodeKey, 2, InStr(1, strNodeKey, "/") - 2)
    End If
End Sub

Private Sub Settings_Click()
    frmSettings.Show vbModal, Me
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
      Case 0
        If blnbuildCategory = False Then
            buildCategory
            blnbuildCategory = True
        End If
        
      Case 1
        If blnbuildCompatibility = False Then
            buildCompatibility
            blnbuildCompatibility = True
        End If
        
     Case 2
        If blnbuildToday = False Then
           buildToday
            blnbuildToday = True
            End If
    End Select
End Sub

Private Sub GoSearch()
  Dim x As Node

    If txtSearchstr.Tag <> txtSearchstr.Text Then
        CmdSearch.Tag = ""
    End If

    txtSearchstr.Tag = txtSearchstr.Text

    If CmdSearch.Tag = "" Then
        If ctlTreeview1.SearchNodes(Trim$(txtSearchstr.Text), x) = J2_Treeview.EnError.Search_String_Found Then
            Call ctlTreeview1.HighlightNode(x)
            CmdSearch.Tag = CLng(x.Index) + 1
          Else
            MsgBox "Not found"
            CmdSearch.Tag = ""
        End If
      Else
        If ctlTreeview1.SearchNodes(Trim$(txtSearchstr.Text), x, CInt(CmdSearch.Tag)) = J2_Treeview.EnError.Search_String_Found Then
            Call ctlTreeview1.HighlightNode(x)
            CmdSearch.Tag = CLng(x.Index) + 1
          Else
            MsgBox "No more found"
            CmdSearch.Tag = ""
        End If
    End If
End Sub


Private Sub buildToday()
  Dim intResult As Integer

    With Me.ctlListview1
        
        '- Verbinding naar de database instellen.
        .Connecttype = "Access2000;"
        .DatabaseName = CStr(gStrDatabaseFilename)
        
        '- Aangeven welke velden en Query voor de Nodes
        .SqlQuery = "SELECT tblPS_PSC.PS_ID, tblPS_PSC.PS_LEVEL, tblPS_PSC.PS_DDSUBMITTED, tblPS_PSC.PS_DDReceived, tblPS_PSC.PS_HTTP, tblPS_PSC.PS_LOCALDIR, tblPS_PSC.PS_DOWNLOADED, tblPS_PSC.PS_TITLE, tblPS_PSC.PS_DESCRIPTION, tblPS_PSC.PS_SERVERFILE From tblPS_PSC WHERE (((tblPS_PSC.PS_DDReceived)>=(SELECT Max(tblCS_CurrentSubscribers.CS_DDDate) AS MaxOfCS_DDDate FROM tblCS_CurrentSubscribers))) order by PS_TITLE"
        
        '''debug.print .SqlQuery
        .fldItem = "PS_ID"
        .fldText = "PS_TITLE"
        
        .IconSet = J2_Treeview.enIcons.VbFiles
        '- Daadwerkelijk opbouwen van de List
        intResult = .Populate
    
    End With

End Sub


Private Sub Search(ByRef strSqlString As String)
  Dim intResult As Integer

    With Me.ctlListview2
        
        '- Verbinding naar de database instellen.
        .Connecttype = "Access2000;"
        .DatabaseName = CStr(gStrDatabaseFilename)
        
        '- Aangeven welke velden en Query voor de Nodes
      ' 'debug.print strSqlString
        .SqlQuery = strSqlString
        
        '''debug.print .SqlQuery
        .fldItem = "PS_ID"
        .fldText = "PS_TITLE"
        
        .IconSet = J2_Treeview.enIcons.VbFiles
        '- Daadwerkelijk opbouwen van de List
        intResult = .Populate
        Select Case .TotaalItems
        Case 0
            Me.lblFound.Caption = "There are no results to display."
        
        Case 1
            Me.lblFound.Caption = "There is 1 result displayed"
        
        Case Else
            Me.lblFound.Caption = "There are  " & .TotaalItems & " results displayed"
        End Select
    End With

End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'if key is {Return Key}
        KeyAscii = 0        'keep from beeping
        cmdSearchNow_Click        ' en start the search
    End If                  'duh

End Sub

Private Sub txtSearchstr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'if key is {Return Key}
        KeyAscii = 0        'keep from beeping
        GoSearch        ' en start the search
    End If                  'duh

End Sub
