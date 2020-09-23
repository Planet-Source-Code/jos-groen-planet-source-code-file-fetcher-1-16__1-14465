VERSION 5.00
Begin VB.UserControl ctlDetails 
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   ScaleHeight     =   6315
   ScaleWidth      =   7560
   ToolboxBitmap   =   "frmDetails.ctx":0000
   Begin VB.Frame frmRecord 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   7575
      Begin VB.Frame Frame7 
         Height          =   495
         Left            =   60
         TabIndex        =   29
         Top             =   5835
         Width           =   7455
         Begin VB.TextBox txtAuthorHref 
            Height          =   285
            Left            =   3795
            TabIndex        =   33
            ToolTipText     =   "Click to visit"
            Top             =   150
            Width           =   3615
         End
         Begin VB.TextBox txtAuthor 
            DataField       =   "AU_NAME"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Left            =   735
            TabIndex        =   30
            Top             =   135
            Width           =   2415
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Author:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   75
            TabIndex        =   32
            Top             =   195
            Width           =   630
         End
         Begin VB.Label lblHREF 
            Caption         =   "Href:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3210
            TabIndex        =   31
            Top             =   195
            Width           =   510
         End
      End
      Begin VB.Frame Frame6 
         Height          =   435
         Left            =   60
         TabIndex        =   17
         Top             =   120
         Width           =   7455
         Begin VB.TextBox txtDownloaded 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4935
            TabIndex        =   28
            Top             =   120
            Width           =   240
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   195
            Left            =   780
            TabIndex        =   27
            Top             =   225
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.TextBox txtLevel 
            DataField       =   "PS_LEVEL"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5910
            TabIndex        =   25
            Top             =   135
            Width           =   1485
         End
         Begin VB.TextBox txtId 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            DataField       =   "PS_ID"
            DataSource      =   "datPrimaryRS"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   345
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   135
            Width           =   570
         End
         Begin VB.TextBox txtSend 
            DataField       =   "PS_DDReceived"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1320
            TabIndex        =   19
            Text            =   "01-01-2000"
            Top             =   135
            Width           =   960
         End
         Begin VB.TextBox txtSubmitted 
            DataField       =   "PS_DDSUBMITTED"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3060
            TabIndex        =   18
            Text            =   "01-01-2000"
            Top             =   135
            Width           =   960
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Level:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   7
            Left            =   5520
            TabIndex        =   26
            Top             =   150
            Width           =   360
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Downloaded:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   5
            Left            =   4080
            TabIndex        =   24
            Top             =   150
            Width           =   810
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Sent:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   3
            Left            =   960
            TabIndex        =   23
            Top             =   150
            Width           =   315
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Id:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   0
            Left            =   60
            TabIndex        =   22
            Top             =   150
            Width           =   135
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Submitted:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   4
            Left            =   2340
            TabIndex        =   21
            Top             =   150
            Width           =   660
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Title:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   60
         TabIndex        =   15
         Top             =   600
         Width           =   7455
         Begin VB.TextBox txtTitle 
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
            Left            =   60
            TabIndex        =   16
            Top             =   180
            Width           =   7350
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   60
         TabIndex        =   13
         Top             =   1320
         Width           =   7455
         Begin VB.TextBox txtDesc 
            DataField       =   "PS_DESCRIPTION"
            DataSource      =   "datPrimaryRS"
            Height          =   2925
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   240
            Width           =   7350
         End
      End
      Begin VB.Frame Frame1 
         Height          =   450
         Left            =   60
         TabIndex        =   8
         Top             =   4560
         Width           =   7455
         Begin VB.CommandButton cmdInternet 
            Height          =   285
            Left            =   6780
            Picture         =   "frmDetails.ctx":0312
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Jump to internet page"
            Top             =   120
            Width           =   315
         End
         Begin VB.TextBox txtUrl 
            DataField       =   "PS_HTTP"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Left            =   735
            TabIndex        =   10
            Top             =   135
            Width           =   6045
         End
         Begin VB.CommandButton cmdDownloadFile 
            Height          =   285
            Left            =   7080
            Picture         =   "frmDetails.ctx":06DF
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Download file"
            Top             =   120
            Width           =   315
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "URL:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   45
            TabIndex        =   12
            Top             =   180
            Width           =   450
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   60
         TabIndex        =   5
         Top             =   4980
         Width           =   7455
         Begin VB.TextBox txtSource 
            DataField       =   "PS_SERVERFILE"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Left            =   720
            TabIndex        =   6
            Top             =   120
            Width           =   6675
         End
         Begin VB.Label lblLServerDIr 
            AutoSize        =   -1  'True
            Caption         =   "Source:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   7
            Top             =   180
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Height          =   435
         Left            =   60
         TabIndex        =   1
         Top             =   5400
         Width           =   7455
         Begin VB.TextBox txtDest 
            DataField       =   "PS_LOCALDIR"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Left            =   720
            TabIndex        =   3
            Top             =   120
            Width           =   6360
         End
         Begin VB.CommandButton cmdOpenFile 
            Height          =   285
            Left            =   7080
            Picture         =   "frmDetails.ctx":0829
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Open the file with his default viewer"
            Top             =   120
            Width           =   315
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Dest:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   60
            TabIndex        =   4
            Top             =   180
            Width           =   465
         End
      End
   End
End
Attribute VB_Name = "ctlDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_Database = ""
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_PsId = 0
'Property Variables:
Dim m_Database As String
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_PsId As Long
'Event Declarations:
Event ButtonClicked(Button As enButtonnr)
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private m_rsMedewerker As ADODB.Recordset
Private mbDataChanged As Boolean
Private colBndMedewerker As New BindingCollection
Private Ado As New J2_ADO.clsADO

Enum enButtonnr
 GotoPage = 1
 DownloadFile = 2
 Openfile = 3
 Hyperlink = 4
End Enum








Private Sub cmdDownloadFile_Click()
    RaiseEvent ButtonClicked(DownloadFile)
End Sub

Private Sub cmdInternet_Click()
    RaiseEvent ButtonClicked(GotoPage)
End Sub

Private Sub cmdOpenFile_Click()
    RaiseEvent ButtonClicked(Openfile)
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
    Ado.J2_Disconnect
    PsId = m_PsId
    Populate
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get PsId() As Long
    PsId = m_PsId
End Property

Public Property Let PsId(ByVal New_PsId As Long)
    m_PsId = New_PsId
    PropertyChanged "PsId"
Populate
End Property




Private Sub Command1_Click()
'm_PsId = 200
'm_Database = "C:\PSC\DATA\PSC2000a.mdb"
'Populate

End Sub

Private Sub txtAuthorHref_Click()
    RaiseEvent ButtonClicked(Hyperlink)

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_PsId = m_def_PsId
    m_Database = m_def_Database
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_PsId = PropBag.ReadProperty("PsId", m_def_PsId)
    m_Database = PropBag.ReadProperty("Database", m_def_Database)
    txtUrl.Text = PropBag.ReadProperty("URL", "")
    txtTitle.Text = PropBag.ReadProperty("Title", "")
    txtUrl.Text = PropBag.ReadProperty("URL", "")
    txtDest.Text = PropBag.ReadProperty("File", "")
    txtAuthorHref.Text = PropBag.ReadProperty("AuthorHref", "")
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    m_rsMedewerker.Close
    Set m_rsMedewerker = Nothing
    Ado.J2_Disconnect
    Set Ado = Nothing
    Set colBndMedewerker = Nothing
    On Error GoTo 0
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("PsId", m_PsId, m_def_PsId)
    Call PropBag.WriteProperty("Database", m_Database, m_def_Database)
    Call PropBag.WriteProperty("URL", txtUrl.Text, "")
    Call PropBag.WriteProperty("Title", txtTitle.Text, "")
    Call PropBag.WriteProperty("URL", txtUrl.Text, "")
    Call PropBag.WriteProperty("File", txtDest.Text, "")
    Call PropBag.WriteProperty("AuthorHref", txtAuthorHref.Text, "")
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Database() As String
    Database = m_Database
End Property

Public Property Let Database(ByVal New_Database As String)
    m_Database = New_Database
    PropertyChanged "Database"
End Property


Public Function Populatex() As Integer

  Dim intResult As Integer
  Dim sqlMedQuery As String
  
     
  Dim rsTemp As ADODB.Recordset


    sqlMedQuery = "SELECT tblPS_PSC.PS_ID, tblPS_PSC.PS_LEVEL, " & _
    "tblPS_PSC.PS_DDSUBMITTED, tblPS_PSC.PS_DDReceived, tblPS_PSC.PS_HTTP, " & _
    "tblPS_PSC.PS_LOCALDIR, " & _
    "tblPS_PSC.PS_DOWNLOADED, tblPS_PSC.PS_TITLE, " & _
    "tblPS_PSC.PS_DESCRIPTION, tblPS_PSC.PS_SERVERFILE, tblPS_PSC.PS_AU_ID, " & _
    "tblAU_AUTHOR.AU_NAME, tblAU_AUTHOR.AU_HREF " & _
    "FROM tblPS_PSC LEFT JOIN tblAU_AUTHOR ON tblPS_PSC.PS_AU_ID = tblAU_AUTHOR.AU_ID WHERE (((tblPS_PSC.PS_ID)=" & CStr(m_PsId) & "))"
    intResult = Ado.J2_Connect(m_Database, Access2000)
    If intResult = J2_ADO.enError.No_Errors Then
        intResult = Ado.J2_Recordset(sqlMedQuery, adOpenDynamic, adLockOptimistic, -1, rsTemp)
        Set m_rsMedewerker = rsTemp
        
        If intResult = J2_ADO.enError.No_Errors Then
          'debug.print m_rsMedewerker("PS_ID").Value
             'Bind the text boxes to the data provider
            Set colBndMedewerker.DataSource = m_rsMedewerker
             ' Add to the Bindings collection.
            
            ''debug.print m_rsMedewerker.Fields("PS_DDReceived").Type
            With colBndMedewerker
                .Clear
                 .Add UserControl.txtId, "Text", "PS_ID", , "PS_ID"
               .Add UserControl.txtSend, "Text", "PS_DDReceived", , "PS_DDReceived"
          .Add UserControl.txtSubmitted, "Text", "PS_DDSUBMITTED", , "PS_DDSUBMITTED"
         .Add UserControl.txtDownloaded, "Text", "PS_DOWNLOADED", , "PS_DOWNLOADED"
              .Add UserControl.txtLevel, "Text", "PS_LEVEL", , "PS_LEVEL"
                 
              .Add UserControl.txtTitle, "Text", "PS_TITLE", , "PS_TITLE"
               
               .Add UserControl.txtDesc, "Text", "PS_DESCRIPTION", , "PS_DESCRIPTION"
               
                .Add UserControl.txtUrl, "Text", "PS_HTTP", , "PS_HTTP"
             .Add UserControl.txtSource, "Text", "PS_SERVERFILE", , "PS_SERVERFILE"
               .Add UserControl.txtDest, "Text", "PS_LOCALDIR", , "PS_LOCALDIR"
             .Add UserControl.txtAuthor, "Text", "AU_NAME", , "AU_NAME"
             .Add UserControl.txtAuthorHref, "Text", "AU_HREF", , "AU_HREF"
       'AU_HREF
       
            End With
            colBndMedewerker.UpdateMode = vbUpdateWhenPropertyChanges
    
            
               
            
           
        End If
     End If
    Dim intTeller As Integer

'    For intTeller = 0 To UserControl.Option.Count - 1
'        If UserControl.Option(intTeller).Value = True Then
'            m_intCurrentOption = intTeller
'        End If
'    Next intTeller
'    Option_Click (m_intCurrentOption)

    Populatex = intResult

End Function



Public Sub Populate()
    Call Populatex
    Exit Sub
    
  Set adoPrimaryRS = New Recordset
    
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & m_Database & ";"

  
  
  '
  
  
  adoPrimaryRS.Open "SELECT tblPS_PSC.PS_ID, tblPS_PSC.PS_LEVEL, tblPS_PSC.PS_DDSUBMITTED, tblPS_PSC.PS_DDReceived, tblPS_PSC.PS_HTTP, tblPS_PSC.PS_LOCALDIR, tblPS_PSC.PS_CODEOFTHEDAY, tblPS_PSC.PS_DOWNLOADED, tblPS_PSC.PS_TITLE, tblPS_PSC.PS_MSGID, tblPS_PSC.PS_DESCRIPTION, tblPS_PSC.PS_SERVERFILE, tblPS_PSC.PS_AU_ID, tblPS_PSC.PS_DELETEDFROMSERVER, tblAU_AUTHOR.AU_NAME FROM tblPS_PSC INNER JOIN tblAU_AUTHOR ON tblPS_PSC.PS_AU_ID = tblAU_AUTHOR.AU_ID WHERE (((tblPS_PSC.PS_ID)=" & CStr(m_PsId) & "))", db, adOpenKeyset, adLockOptimistic
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  Dim oCheck As CheckBox

  mbDataChanged = False
  adoPrimaryRS.Close
  db.Close
  
End Sub
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=13,0,0,
'Public Property Get URL() As String
'    URL = m_URL
'End Property
'
'Public Property Let URL(ByVal New_URL As String)
'    m_URL = New_URL
'    PropertyChanged "URL"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtTitle,txtTitle,-1,Text
Public Property Get Title() As String
Attribute Title.VB_Description = "Returns/sets the text contained in the control."
    Title = txtTitle.Text
End Property

Public Property Let Title(ByVal New_Title As String)
    txtTitle.Text() = New_Title
    PropertyChanged "Title"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtUrl,txtUrl,-1,Text
Public Property Get URL() As String
Attribute URL.VB_Description = "Returns/sets the text contained in the control."
    URL = txtUrl.Text
End Property

Public Property Let URL(ByVal New_URL As String)
    txtUrl.Text() = New_URL
    PropertyChanged "URL"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDest,txtDest,-1,Text
Public Property Get File() As String
Attribute File.VB_Description = "Returns/sets the text contained in the control."
    File = txtDest.Text
End Property

Public Property Let File(ByVal New_File As String)
    txtDest.Text() = New_File
    PropertyChanged "File"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtAuthorHref,txtAuthorHref,-1,Text
Public Property Get AuthorHref() As String
Attribute AuthorHref.VB_Description = "Returns/sets the text contained in the control."
    AuthorHref = txtAuthorHref.Text
End Property

Public Property Let AuthorHref(ByVal New_AuthorHref As String)
    txtAuthorHref.Text() = New_AuthorHref
    PropertyChanged "AuthorHref"
End Property

