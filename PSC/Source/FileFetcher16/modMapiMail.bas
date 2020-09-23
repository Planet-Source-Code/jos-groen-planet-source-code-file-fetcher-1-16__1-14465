Attribute VB_Name = "modMapiMail"

Option Explicit

Public Type Outlook
    Session As MAPISession
    Messages As MAPIMessages
    lFoundMails As Long             'Hoeveel berichten gevonden die voldoen
    aMsgIndexArray() As Long        'Msgindex opslaan van berichten die we willen analyseren
    lFoundSubmissions As Long ' Aantal Subscritions die zijn gevonden
End Type

Dim Uitkijk As Outlook

Function Uitkijk_Logon(Session As MAPISession, Messages As MAPIMessages) As Integer

    On Error GoTo Uitkijk_Logon_Err
    'Set Uitkijk = Outlook

    Set Uitkijk.Session = Session
    Set Uitkijk.Messages = Messages

    DisplayInfo "Signing in..."
    Uitkijk.Session.DownLoadMail = False
    Uitkijk.Session.SignOn
   
    Uitkijk.Messages.SessionID = Uitkijk.Session.SessionID

    DisplayInfo "Fetching messages..."
    Uitkijk.Messages.Fetch
    DisplayInfo "Fetched " & Uitkijk.Messages.MsgCount & " messages"

Exit Function

Uitkijk_Logon_Err:
    
    If Err.Number = 32050 Then  'Logon failure: valid session ID already exists
        
        DisplayInfo ("[" & Err.Description & "]" & Uitkijk.Session.SessionID)
        Resume Next
      Else
        DisplayInfo ("[" & Err.Number & "-" & Err.Description & "]")
        Exit Function
    End If

End Function

Function Uitkijk_Analyse() As Boolean
  
  On Error GoTo Uitkijk_Analyse_Error
  Dim lMsgIndex As Long ' lus tellertje om door alle mailtjes te browse
  Dim VarReturnValue As Variant '
  Dim intResult As Integer
    '  On Error GoTo Uitkijk_Analyse_err
    Uitkijk.lFoundMails = 0
    Uitkijk.Messages.FetchSorted = True
    For lMsgIndex = 0 To Uitkijk.Messages.MsgCount - 1
        Call DisplayInfoProgressbar1(lMsgIndex, Uitkijk.Messages.MsgCount - 1)
        Uitkijk.Messages.MsgIndex = lMsgIndex
        DisplayInfoLbl "Analyzing message: " & Uitkijk.Messages.MsgIndex + 1 & " of " & Uitkijk.Messages.MsgCount
        If Uitkijk.Messages.MsgOrigAddress = "MailingList@planet-source-code.com" Or _
           Uitkijk.Messages.MsgOrigAddress = "IanIppolito@planet-source-code.com" Then
            '      And Uitkijk.Messages.MsgDateReceived > Date - 1 Then
            If Format$(Uitkijk.Messages.MsgDateReceived, "dd/mm/yyyy") >= g_datScan Then
                Uitkijk.lFoundMails = Uitkijk.lFoundMails + 1
                ReDim Preserve Uitkijk.aMsgIndexArray(Uitkijk.lFoundMails)
                Uitkijk.aMsgIndexArray(Uitkijk.lFoundMails) = Uitkijk.Messages.MsgIndex
                'debug.print "ScanedDate:" & Format$(Uitkijk.Messages.MsgDateReceived, "dd/mm/yyyy")
                
                'if auto update last scanned mail then update that value
                 VarReturnValue = 0
                 intResult = Reg.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "AutoUpdateDate", VarReturnValue)
                 If VarReturnValue = 1 Then
                    intResult = Reg.SetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & g_ApplTitle & "\Options", "Scandate", CStr(Uitkijk.Messages.MsgDateReceived))
                 End If
                
            End If
        End If
    
    Next lMsgIndex
    Uitkijk_Analyse = True

    If Uitkijk.lFoundMails = 0 Then
        DisplayInfoLbl "No messages found."
      Else
        DisplayInfoLbl Uitkijk.lFoundMails & " messages found"
        lMsgIndex = 0
        Uitkijk.lFoundSubmissions = 0
        For lMsgIndex = 1 To Uitkijk.lFoundMails
            Call DisplayInfoProgressbar1(lMsgIndex, Uitkijk.lFoundMails + 1)
            DisplayInfo CStr(Uitkijk.aMsgIndexArray(lMsgIndex))
            Call Uitkijk_Analyse_Message(Uitkijk.aMsgIndexArray(lMsgIndex))
        Next lMsgIndex
        Call DisplayInfoProgressbar1(lMsgIndex, Uitkijk.lFoundMails + 1)
        ReDim Uitkijk.aMsgIndexArray(0)
        DisplayInfoLbl "Done... " & Uitkijk.lFoundSubmissions & " Submissions found"
    End If
    DisplayInfoProgressbar1Hide

Exit Function
Uitkijk_Analyse_Error:
Uitkijk_Analyse_err:
    DisplayInfoLbl "Done with errors ... " & Uitkijk.lFoundMails & " messages found"

End Function

Sub Uitkijk_logoff()

    On Error Resume Next
      If Uitkijk.Session.SessionID > 0 Then
          Uitkijk.Session.SignOff
      End If
      Set Uitkijk.Session = Nothing
      Set Uitkijk.Messages = Nothing
    On Error GoTo 0

End Sub

Function Uitkijk_Analyse_Message(lMsgIndex As Long) As Boolean

  Dim sTemp As String
  Dim lLusteller As String        'Zoek alle subscibes

    On Error GoTo Uitkijk_Analyse_Message_Err
    lLusteller = 1
    
    Uitkijk.Messages.MsgIndex = lMsgIndex
    ' === GET : Current # of subscribers:
    DATA_UpdateCurrentSubscribers CLng(STRING_GetStringBetween(Uitkijk.Messages.MsgNoteText, "Current # of subscribers:", vbCr & vbLf)), Uitkijk.Messages.MsgDateReceived
    
    Do While Not InStr(1, Uitkijk.Messages.MsgNoteText, vbCr & vbLf & CStr(lLusteller) & ")", vbTextCompare) And IsNull(InStr(1, Uitkijk.Messages.MsgNoteText, vbCr & vbLf & CStr(lLusteller) & ")", vbTextCompare)) = False
        'DisplayInfoCls
        DisplayInfoLbl1 "Analyse subscribe:" & CStr(lLusteller)
        sTemp = STRING_GetStringBetween(Uitkijk.Messages.MsgNoteText, vbCr & vbLf & CStr(lLusteller) & ")", vbCr & vbLf)
        
        'Title
        sTemp = STRING_GetStringBetween(Uitkijk.Messages.MsgNoteText, vbCr & vbLf & CStr(lLusteller) & ")", vbCr & vbLf, StrPosAfterSearch)
        sTemp = Replace$(sTemp, vbCr & vbLf, " ")
        sTemp = Replace$(sTemp, """", "'")
        Subscriber.Title = Trim$(sTemp)
        DisplayInfo STRING_Uitlijnen("Title:") & Trim$(sTemp)
        
        'Category:"
        sTemp = STRING_GetStringBetween(Uitkijk.Messages.MsgNoteText, vbCr & vbLf & "Category:", "Lev", StrPosAfterSearch)
        sTemp = Replace$(sTemp, vbLf, " ")
        sTemp = Replace$(sTemp, vbCr, " ")
        sTemp = Replace$(sTemp, vbCr & vbLf, " ")
        Uitkijk_Analyse_Category (Trim$(sTemp))
        
        'Level:
        sTemp = STRING_GetStringBetween(Uitkijk.Messages.MsgNoteText, "Level:", vbCr & vbLf, StrPosAfterSearch)
        sTemp = Replace$(sTemp, vbLf, " ")
        sTemp = Replace$(sTemp, vbCr, " ")
        sTemp = Replace$(sTemp, vbCr & vbLf, " ")
        Subscriber.Level = Trim$(sTemp)
        DisplayInfo STRING_Uitlijnen("Level:") & Trim$(sTemp)
        
        'Description:
        sTemp = STRING_GetStringBetween(Uitkijk.Messages.MsgNoteText, vbCr & vbLf & "Description:", vbCr & vbLf & vbCr & vbLf, StrPosAfterSearch)
        sTemp = Replace$(sTemp, vbLf, " ")
        sTemp = Replace$(sTemp, vbCr, " ")
        sTemp = Replace$(sTemp, vbCr & vbLf, " ")
        Subscriber.Description = Trim$(sTemp)
        DisplayInfo STRING_Uitlijnen("Description:") & Trim$(sTemp)
        
        'Complete source code is at:
        sTemp = STRING_GetStringBetween(Uitkijk.Messages.MsgNoteText, vbCr & vbLf & "Complete source code is at:", "Compa", StrPosAfterSearch)
        sTemp = Replace$(sTemp, vbLf, " ")
        sTemp = Replace$(sTemp, vbCr, " ")
        Subscriber.source_code_at = Trim$(sTemp)

        DisplayInfo STRING_Uitlijnen("source code at:") & Trim$(sTemp)
        
        'Compatibility:
        sTemp = STRING_GetStringBetween(Uitkijk.Messages.MsgNoteText, "Compatibility:", "Sub", StrPosAfterSearch)
        sTemp = Replace$(sTemp, vbLf, " ")
        sTemp = Replace$(sTemp, vbCr, " ")
        Uitkijk_Analyse_Compatibility (Trim$(sTemp))
        
        'Submitted on
        sTemp = STRING_GetStringBetween(Uitkijk.Messages.MsgNoteText, "Submitted on", "and accessed", StrPosAfterSearch)
        Subscriber.Submitted_on = Format$(Trim$(sTemp), "mm-dd-yyyy")
        DisplayInfo STRING_Uitlijnen("Submitted on:") & Trim$(sTemp)
        
        'Datereceived
        Subscriber.DateReceived = Uitkijk.Messages.MsgDateReceived
        DisplayInfo STRING_Uitlijnen("PS_ID:") & CStr(DATA_UpdatePSC)
        lLusteller = lLusteller + 1
        If InStr(1, Uitkijk.Messages.MsgNoteText, CStr(lLusteller) & ")", vbTextCompare) = 0 Then
            Exit Do
        End If
    Loop
    lLusteller = lLusteller - 2
    Uitkijk.lFoundSubmissions = Uitkijk.lFoundSubmissions + lLusteller
    DisplayInfoLbl1 "Done... Submissions found:" & lLusteller

Exit Function

Uitkijk_Analyse_Message_Err:
    DisplayInfoLbl1 "Done With Errors... Submissions found::" & lLusteller

End Function

Sub Uitkijk_Analyse_Compatibility(strCompatibility As String)

  Dim sTemp As String
  Dim iAantalArgs As Integer
  Dim ibrowsecompa As Integer
  Dim lCnId As Long           'compatibity names id

    sTemp = Replace$(strCompatibility, ",", ";")
    sTemp = Replace$(sTemp, "  ", " ")

    iAantalArgs = STRING_CountArgs(sTemp)
   
    'Opslaan in een array
    ReDim g_vCompatibility(iAantalArgs)
    g_vCompatibilityCount = iAantalArgs
    
    DisplayInfo STRING_Uitlijnen("- Compatibility:")
    For ibrowsecompa = 1 To iAantalArgs
        lCnId = DATA_UpdateCompatibility(Trim$(STRING_GeefArg(ibrowsecompa, sTemp)))
        g_vCompatibility(ibrowsecompa) = lCnId
        DisplayInfo STRING_Uitlijnen("-         :") & "[" & lCnId & "] " & Trim$(STRING_GeefArg(ibrowsecompa, sTemp))
    Next ibrowsecompa

End Sub

Sub Uitkijk_Analyse_Category(strCategory As String)

  Dim sTemp As String
  Dim iAantalArgs As Integer
  Dim ibrowsecompa As Integer
  Dim lCnId As Long           'compatibity names id

    sTemp = Replace$(strCategory, "/", ";")
    sTemp = Replace$(sTemp, "  ", " ")
    
    iAantalArgs = STRING_CountArgs(sTemp)
   
    'Opslaan in een array
    ReDim g_vCategories(iAantalArgs)
    g_vCategoriesCount = iAantalArgs
     
    DisplayInfo STRING_Uitlijnen("- Categories:")
    For ibrowsecompa = 1 To iAantalArgs
        
        lCnId = DATA_UpdateCategory(Trim$(STRING_GeefArg(ibrowsecompa, sTemp)))
        'Fill Array
        g_vCategories(ibrowsecompa) = lCnId
        DisplayInfo STRING_Uitlijnen("-         :") & "[" & lCnId & "] " & Trim$(STRING_GeefArg(ibrowsecompa, sTemp))
        
    Next ibrowsecompa

End Sub
