VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlListview 
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3705
   ScaleHeight     =   4830
   ScaleWidth      =   3705
   ToolboxBitmap   =   "ctlListview.ctx":0000
   Begin MSComctlLib.ListView LstView 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   7435
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "imglstTreeview"
      SmallIcons      =   "imglstTreeview"
      ColHdrIcons     =   "imglstTreeview"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imglstTreeview 
      Left            =   2880
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":0312
            Key             =   "keyMap"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":08AC
            Key             =   "keyMedewerker"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":1188
            Key             =   "keyMedewerkerDetails"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":1A64
            Key             =   "keyGroep"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":2340
            Key             =   "keyGroepDetails"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":2C1C
            Key             =   "keyMapOpen"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":31B6
            Key             =   "cursor"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":34D0
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":3A6C
            Key             =   "keyAb"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":3BC8
            Key             =   "keyBook"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":4E4C
            Key             =   "keyBooks"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":52A0
            Key             =   "keyBooksOpen"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":56F4
            Key             =   "keyWatch"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":5B48
            Key             =   "keyMannetje"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":60E4
            Key             =   "keyMannetjes"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":69C0
            Key             =   "keyMannetjerood"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":6F5C
            Key             =   "keyMapdicht2"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":7838
            Key             =   "keyMapOpen2"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlListview.ctx":8114
            Key             =   "keyvbprj"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ctlListview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'bug fix list
'1.08 returns the number of added items
'1.07 fixed ColumnHeader

Const m_def_IconSet = 0 ':( As Integer ?


Dim m_DatabaseName As String
Dim m_Connecttype As String
Dim m_SqlQuery As String
Dim m_fldText As String
Dim m_IconSet As Integer

Dim m_rsListview As adodb.Recordset     ' Door de hele Control heen wordt deze recordset gebruikt.
Dim Ado As New J2_ADO.clsADO           ' Verbinding naar de Ado Dll.

'/======<ICONENSETS>================================
Public Enum enIcons
    Default = 0
    Maps = 1
    Books = 2
    Tasks = 3
    Persons = 4
    Organisation = 5
    VbFiles = 6
End Enum

Private m_stricoMap        As String
Private m_stricoMapOpen   As String
Private m_stricoChild     As String
Private m_stricoChildOpen As String

'Default Property Values:
Const m_def_TotaalItems = 0
'Const m_def_TotaalItems = 0
Const m_def_fldItem = ""
'Const m_def_AutoRedraw = 0
'Const m_def_BackColor = 0
'Const m_def_BackStyle = 0

'Property Variables:
Dim m_TotaalItems As Long
'Dim m_TotaalItems As Long
Dim m_fldItem As String
'Dim m_AutoRedraw As Boolean
'Dim m_BackColor As Long
'Dim m_BackStyle As Integer
'Event Declarations:
Event Click(ByVal Item As ListItem) 'MappingInfo=LstView,LstView,-1,ItemClick
Attribute Click.VB_Description = "Occurs when a ListItem object is clicked or selected"





Private Sub SelectIconset(ByRef IconSet As Integer)

    Select Case IconSet

      Case enIcons.Maps, enIcons.Default
        m_stricoMap = "keyMap"
        m_stricoMapOpen = "keyMapOpen"
        m_stricoChild = "keyAb"
        m_stricoChildOpen = "keyAb"

      Case enIcons.Books
        m_stricoMap = "keyBooks"
        m_stricoMapOpen = "keyBooksOpen"
        m_stricoChild = "keyBook"
        m_stricoChildOpen = "keyBook"
    
      Case enIcons.Tasks
        m_stricoMap = "keyMapdicht2"
        m_stricoMapOpen = "keyMapOpen2"
        m_stricoChild = "keyWatch"
        m_stricoChildOpen = "keyWatch"

      Case enIcons.Persons
        m_stricoMap = "keyMannetjes"
        m_stricoMapOpen = "keyMannetjerood"
        m_stricoChild = "keyMannetje"
        m_stricoChildOpen = "keyMannetjerood"

      Case enIcons.Organisation
        m_stricoMap = "keyGroepDetails"
        m_stricoMapOpen = "keyGroep"
        m_stricoChild = "keyMedewerker"
        m_stricoChildOpen = "keyMedewerkerDetails"

      Case enIcons.VbFiles
        m_stricoMap = "keyMapdicht2"
        m_stricoMapOpen = "keyMapOpen2"
        m_stricoChild = "keyvbprj"
        m_stricoChildOpen = "keyvbprj"
    
      Case Else
        m_stricoMap = "keyMap"
        m_stricoMapOpen = "keyMapOpen"
        m_stricoChild = "keyAb"
        m_stricoChildOpen = "keyAb"
        
    End Select

End Sub

'/======<ICONENSETS END >================================



Public Function Populate() As Integer
  Dim IntResult As Integer
    IntResult = Ado.J2_Connect(m_DatabaseName, Access2000)
    If IntResult Then
        IntResult = Ado.J2_Recordset(m_SqlQuery, adOpenKeyset, adLockOptimistic, -1, m_rsListview)
        If IntResult = J2_ADO.enError.No_Errors Then
            populatelist
            m_rsListview.Close
            Set m_rsListview = Nothing
        End If
        Ado.J2_Disconnect
    End If
    Populate = IntResult
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Databasename() As String

    Databasename = m_DatabaseName

End Property

Public Property Let Databasename(ByVal New_DatabaseName As String)

    m_DatabaseName = New_DatabaseName
    PropertyChanged "Databasename"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Connecttype() As String

    Connecttype = m_Connecttype

End Property

Public Property Let Connecttype(ByVal New_Connecttype As String)

    m_Connecttype = New_Connecttype
    PropertyChanged "Connecttype"

End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get SqlQuery() As String

    SqlQuery = m_SqlQuery

End Property

Public Property Let SqlQuery(ByVal New_SqlQuery As String)

    m_SqlQuery = New_SqlQuery
    PropertyChanged "sqlQuery"

End Property


Sub populatelist()
        
        Dim clmAdd As ColumnHeader
        Dim itmAdd As ListItem
   
        UserControl.LstView.ListItems.Clear
        UserControl.LstView.ColumnHeaders.Clear
        m_TotaalItems = 0
        'Add two Column Headers to the ListView control
        Set clmAdd = LstView.ColumnHeaders.Add(, , "Title", UserControl.LstView.Width - 40)
   
        'Set the view property of the Listview control to Report view
        LstView.View = lvwReport
        
        
        Do Until m_rsListview.EOF
            'Add data to the ListView control
            Set itmAdd = LstView.ListItems.Add(, "I" & CStr(m_rsListview(m_fldItem).Value), CStr(m_rsListview(m_fldText).Value), , m_stricoChild) ' Author.
            m_TotaalItems = m_TotaalItems + 1
            m_rsListview.MoveNext
        Loop
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get fldText() As String

    fldText = m_fldText

End Property

Public Property Let fldText(ByVal New_fldText As String)

    m_fldText = New_fldText
    PropertyChanged "fldText"

End Property

Private Sub LstView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RaiseEvent Click(Item)

End Sub

Private Sub UserControl_Resize()
    LstView.Left = 0
    LstView.Top = 0
    LstView.Width = ScaleWidth
    LstView.Height = ScaleHeight
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get IconSet() As Integer

    IconSet = m_IconSet

End Property

Public Property Let IconSet(ByVal New_IconSet As Integer)

    m_IconSet = New_IconSet
    PropertyChanged "IconSet"
    SelectIconset m_IconSet

End Property
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=0,0,0,0
'Public Property Get AutoRedraw() As Boolean
'    AutoRedraw = m_AutoRedraw
'End Property
'
'Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
'    m_AutoRedraw = New_AutoRedraw
'    PropertyChanged "AutoRedraw"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=8,0,0,0
'Public Property Get BackColor() As Long
'    BackColor = m_BackColor
'End Property
'
'Public Property Let BackColor(ByVal New_BackColor As Long)
'    m_BackColor = New_BackColor
'    PropertyChanged "BackColor"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=7,0,0,0
'Public Property Get BackStyle() As Integer
'    BackStyle = m_BackStyle
'End Property
'
'Public Property Let BackStyle(ByVal New_BackStyle As Integer)
'    m_BackStyle = New_BackStyle
'    PropertyChanged "BackStyle"
'End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_AutoRedraw = m_def_AutoRedraw
'    m_BackColor = m_def_BackColor
'    m_BackStyle = m_def_BackStyle
    m_fldItem = m_def_fldItem
'    m_TotaalItems = m_def_TotaalItems
    m_TotaalItems = m_def_TotaalItems
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    m_AutoRedraw = PropBag.ReadProperty("AutoRedraw", m_def_AutoRedraw)
'    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
'    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_fldItem = PropBag.ReadProperty("fldItem", m_def_fldItem)
    LstView.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    LstView.Appearance = PropBag.ReadProperty("Appearance", 1)
'    m_TotaalItems = PropBag.ReadProperty("TotaalItems", m_def_TotaalItems)
    m_TotaalItems = PropBag.ReadProperty("TotaalItems", m_def_TotaalItems)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("AutoRedraw", m_AutoRedraw, m_def_AutoRedraw)
'    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
'    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("fldItem", m_fldItem, m_def_fldItem)
    Call PropBag.WriteProperty("BackColor", LstView.BackColor, &H80000005)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("Appearance", LstView.Appearance, 1)
'    Call PropBag.WriteProperty("TotaalItems", m_TotaalItems, m_def_TotaalItems)
    Call PropBag.WriteProperty("TotaalItems", m_TotaalItems, m_def_TotaalItems)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get fldItem() As String
    fldItem = m_fldItem
End Property

Public Property Let fldItem(ByVal New_fldItem As String)
    m_fldItem = New_fldItem
    PropertyChanged "fldItem"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LstView,LstView,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = LstView.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    LstView.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LstView,LstView,-1,Appearance
Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not controls, Forms or an MDIForm are painted at run time with 3-D effects."
    Appearance = LstView.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    LstView.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=8,1,2,0
'Public Property Get TotaalItems() As Long
'    TotaalItems = m_TotaalItems
'End Property
'
'Public Property Let TotaalItems(ByVal New_TotaalItems As Long)
'    If Ambient.UserMode = False Then Err.Raise 387
'    If Ambient.UserMode Then Err.Raise 382
'    m_TotaalItems = New_TotaalItems
'    PropertyChanged "TotaalItems"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,1,1,0
Public Property Get TotaalItems() As Long
    TotaalItems = m_TotaalItems
End Property

Public Property Let TotaalItems(ByVal New_TotaalItems As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_TotaalItems = New_TotaalItems
    PropertyChanged "TotaalItems"
End Property

