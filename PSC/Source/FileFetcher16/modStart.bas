Attribute VB_Name = "modStart"

Option Explicit

Public Type Subscribe
    Title           As String
    Level           As String
    Description     As String
    source_code_at  As String
    Submitted_on    As Date
    DateReceived    As Date
End Type
Global Subscriber As Subscribe
Global File As New J2_FILE.clsFile

Global Ado As New J2_ADO.clsADO
Global Reg As New J2_Registry.clsRegistry
Global Dte As New J2_Date.Date
Global gStrDatabaseFilename As String
Global gStrDefaultDownloadDir As String

Global g_bFeedback As Integer          'Displayinfo functies aan bij ja

Public g_vCategories As Variant         'Hier worden tijdelijk de categories opgeslagen
Global g_vCategoriesCount As Long      'Hier worden het aantal opgeslagen
Public g_vCompatibility As Variant         'Hier worden tijdelijk de Compatibility opgeslagen
Global g_vCompatibilityCount As Long      'Hier worden het aantal opgeslagen
Global g_ApplTitle As String              'Wordt gevuld voor Caption title form
Global g_datScan As Date                  'the date when the scan starts
Global g_Logfile As String                 'The fullpath of the logfile


Public g_vSearch As Variant         'Hier worden tijdelijk de categories opgeslagen
Global g_vSearchCount As Long      'Hier worden het aantal opgeslagen

Sub Main()
g_ApplTitle = App.Title & "  " & App.Major & "." & App.Minor & App.Revision
  Dim iResult As Integer
    Dim intResult As Integer
    Dim VarReturnValue As Variant
    
    VarReturnValue = 0
    intResult = Reg.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "Scandate", VarReturnValue)
    g_datScan = VarReturnValue
   intResult = Reg.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "DefaultDownloadDir", VarReturnValue)
    gStrDefaultDownloadDir = VarReturnValue
   
   intResult = Reg.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "DatabaseLocation", VarReturnValue)
    gStrDatabaseFilename = VarReturnValue

       intResult = Reg.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "Logfile", VarReturnValue)
    g_Logfile = IIf(VarReturnValue = 0, App.Path & "\" & "PSCLOGFILE.TXT", VarReturnValue)
    

    If File.FileExists(gStrDatabaseFilename) = -1 Then
        If File.Exists(App.Path & "PSC2000a.mdb") = -1 Then
            MsgBox "database not found please edit your settings! and restart the application:" & gStrDatabaseFilename
        End If
        
    End If
    
    DisplayInfoCls
    DisplayInfoNOcrlf ("Connect to database :" & gStrDatabaseFilename)
    
    If File.FileExists(gStrDatabaseFilename) Then
        
        iResult = Ado.J2_Connect(gStrDatabaseFilename, Access2000)
        
        If iResult = J2_ADO.EnError.No_Errors Then
            DisplayInfoNOcrlf (" [Established]")
          Else
            DisplayInfoNOcrlf (" [not Connected] Errnr:" & iResult)
        End If
      
      Else
        DisplayInfoNOcrlf " [database not found] select settings to modify"
    End If
    
    Set File = Nothing


    Load Form1
    Form1.Show
End Sub


Sub Reconnnect()
        Dim iResult As Integer
        iResult = Ado.J2_Connect(gStrDatabaseFilename, Access2000)
        
        If iResult = J2_ADO.EnError.No_Errors Then
        Else
        MsgBox "Connection to database Failed" & vbLf & Ado.J2_ErrorMsg(iResult)
        End If
        
End Sub
