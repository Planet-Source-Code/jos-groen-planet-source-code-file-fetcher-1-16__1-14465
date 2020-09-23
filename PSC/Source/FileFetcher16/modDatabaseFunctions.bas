Attribute VB_Name = "modDatabaseFunctions"
Option Explicit

Sub DATA_UpdateCurrentSubscribers(lngCurrentSubscribers As Long, ddDate As Date)

  Dim iResult As Integer
  Dim rstCurrentSubscribers As adodb.Recordset
  Dim strSql As String

    strSql = "SELECT * FROM tblCS_CurrentSubscribers " & _
             " WHERE" & _
             " CS_CurrentSubscribers =" & lngCurrentSubscribers & _
             " AND CS_DDDate =" & "#" & Format$(ddDate, "mm-dd-yyyy") & "#" & _
             ""
    iResult = Ado.J2_Recordset(strSql, adOpenDynamic, adLockOptimistic, adCmdText, rstCurrentSubscribers)
    If iResult = J2_ADO.EnError.No_Errors Then
        If (rstCurrentSubscribers.EOF And rstCurrentSubscribers.BOF) Then
            rstCurrentSubscribers.AddNew
            rstCurrentSubscribers("CS_CurrentSubscribers") = lngCurrentSubscribers
            rstCurrentSubscribers("CS_DDDate") = Format$(ddDate, "dd-mm-yyyy")
            rstCurrentSubscribers.Update
        End If
      Else
        DisplayInfo ("Update CurrentSubscribers failed")
    End If
    rstCurrentSubscribers.Close
    Set rstCurrentSubscribers = Nothing

End Sub

'Result is a CN_ID nummertje
Function DATA_UpdateCompatibility(strCompatibility As String) As Long

  Dim iResult As Integer
  Dim rstCom As adodb.Recordset
  Dim strSql As String
  
    strSql = "SELECT CN_ID,CN_CompatibilityName FROM tblCN_CompatabilityNames" & _
             " WHERE" & _
             " CN_CompatibilityName =" & "'" & strCompatibility & "'"
    iResult = Ado.J2_Recordset(strSql, adOpenDynamic, adLockOptimistic, adCmdText, rstCom)
    
    If iResult = J2_ADO.EnError.No_Errors Then
        If (rstCom.EOF And rstCom.BOF) Then
            rstCom.AddNew
            rstCom("CN_CompatibilityName") = strCompatibility
            DATA_UpdateCompatibility = rstCom("CN_ID")
            rstCom.Update
          Else
            DATA_UpdateCompatibility = rstCom("CN_ID")
        End If
      Else
        DisplayInfo ("Update Compatibility failed")
    End If
    rstCom.Close
    Set rstCom = Nothing

End Function

'Result is a CN_ID nummertje
Function DATA_UpdateCategory(strCategory As String) As Long

  Dim iResult As Integer
  Dim rstCom As adodb.Recordset
  Dim strSql As String
  
    strSql = "SELECT CA_ID,CA_CategoryName FROM tblCA_CategoryNames" & _
             " WHERE" & _
             " CA_CategoryName =" & "'" & strCategory & "'"
    iResult = Ado.J2_Recordset(strSql, adOpenDynamic, adLockOptimistic, adCmdText, rstCom)
    
    If iResult = J2_ADO.EnError.No_Errors Then
        If (rstCom.EOF And rstCom.BOF) Then
            rstCom.AddNew
            rstCom("CA_CategoryName") = strCategory
            DATA_UpdateCategory = rstCom("CA_ID")
            rstCom.Update
          Else
            DATA_UpdateCategory = rstCom("CA_ID")
        End If
      Else
        '''''debug.print strSql
        DisplayInfo ("Update Category failed")
    End If
    On Error Resume Next
      rstCom.Close
      Set rstCom = Nothing
    On Error GoTo 0

End Function

Function DATA_UpdatePSC() As Long

  Dim iResult As Integer
  Dim rstPSC As adodb.Recordset
  Dim strSql As String
  
    strSql = "SELECT PS_ID,  PS_TITLE, PS_LEVEL, PS_DESCRIPTION,  PS_DDSUBMITTED, PS_DDReceived, PS_HTTP, PS_LOCALDIR, PS_DOWNLOADED From tblPS_PSC" & _
             " WHERE PS_TITLE =" & """" & Subscriber.Title & """"
    
    iResult = Ado.J2_Recordset(strSql, adOpenDynamic, adLockOptimistic, adCmdText, rstPSC)
    
    If iResult = J2_ADO.EnError.No_Errors Then
        If (rstPSC.EOF And rstPSC.BOF) Then
            rstPSC.AddNew
            With Subscriber
                rstPSC("PS_DESCRIPTION") = .Description
                rstPSC("PS_LEVEL") = .Level
                rstPSC("PS_HTTP") = .source_code_at
                rstPSC("PS_DDSUBMITTED") = .Submitted_on
                rstPSC("PS_TITLE") = .Title
                rstPSC("PS_DDReceived") = .DateReceived
            End With
            DATA_UpdatePSC = rstPSC("PS_ID")
            rstPSC.Update
            DATA_UpdateCategoryWith rstPSC("PS_ID")
            DATA_UpdateCompatabilityWith rstPSC("PS_ID")
          Else
            If rstPSC("PS_DDReceived") = Subscriber.DateReceived Then
                DATA_UpdatePSC = rstPSC("PS_ID")
              Else
                '''''debug.print "DATA_UpdatePSC: Reeds geimporteerd op" & rstPSC("PS_DDReceived") & " DateReceived:" & Subscriber.DateReceived
            End If
        End If
      Else
        DisplayInfo ("Update PSC failed")
    End If
    
    On Error Resume Next
      rstPSC.Close
      Set rstPSC = Nothing
    On Error GoTo 0

End Function

Sub DATA_UpdateCategoryWith(lPS_ID As Long)

  Dim ltmp As Long
  Dim iResult As Integer
  Dim rstCategoryWith As adodb.Recordset
  Dim strSql As String
  
    strSql = "SELECT CC_CA_ID, CC_PS_ID From tblCC_CategoryWith"
    
    iResult = Ado.J2_Recordset(strSql, adOpenDynamic, adLockOptimistic, adCmdText, rstCategoryWith)
    
    If iResult = J2_ADO.EnError.No_Errors Then
        For ltmp = 1 To g_vCategoriesCount
            rstCategoryWith.AddNew
            With Subscriber
                rstCategoryWith("CC_CA_ID") = CLng(g_vCategories(ltmp))
                rstCategoryWith("CC_PS_ID") = lPS_ID
            End With
            rstCategoryWith.Update
        Next ltmp
      Else
        DisplayInfo ("Update CategoryWith failed")
    End If
    
    On Error Resume Next
      rstCategoryWith.Close
      Set rstCategoryWith = Nothing
    On Error GoTo 0

End Sub

Sub DATA_UpdateCompatabilityWith(lPS_ID As Long)

  Dim ltmp As Long
  Dim iResult As Integer
  Dim rstCompatabilityWith As adodb.Recordset
  Dim strSql As String
  
    strSql = "SELECT CW_PS_ID,CW_CN_ID FROM tblCW_CompatabilityWith"
    
    iResult = Ado.J2_Recordset(strSql, adOpenDynamic, adLockOptimistic, adCmdText, rstCompatabilityWith)
    
    If iResult = J2_ADO.EnError.No_Errors Then
        For ltmp = 1 To g_vCompatibilityCount
            rstCompatabilityWith.AddNew
            With Subscriber
                rstCompatabilityWith("CW_CN_ID") = CLng(g_vCompatibility(ltmp))
                rstCompatabilityWith("CW_PS_ID") = lPS_ID
            End With
            rstCompatabilityWith.Update
        Next ltmp
      Else
        DisplayInfo ("Update CompatabilityWith failed")
    End If
    
    On Error Resume Next
      rstCompatabilityWith.Close
      Set rstCompatabilityWith = Nothing
    On Error GoTo 0

End Sub

'Result is a CN_ID nummertje
Function DATA_UpdateAuthor(strAuthor As String, ByRef strAuthorHref As String) As Long
    If IsNull(Trim(strAuthor)) = True Then
    strAuthor = "Anonymous"
    Else
    strAuthor = Trim(StringStripper(strAuthor))
    End If

  Dim iResult As Integer
  Dim rstCom As adodb.Recordset
  Dim strSql As String
  
    strSql = "SELECT AU_ID,AU_NAME,AU_Href FROM tblAU_AUTHOR" & _
             " WHERE" & _
             " AU_NAME =" & "'" & strAuthor & "'"
    iResult = Ado.J2_Connect(gStrDatabaseFilename, Access2000)
    If iResult = J2_ADO.EnError.No_Errors Then
    
    iResult = Ado.J2_Recordset(strSql, adOpenDynamic, adLockOptimistic, adCmdText, rstCom)
    If iResult = J2_ADO.EnError.No_Errors Then
        If (rstCom.EOF And rstCom.BOF) Then
            rstCom.AddNew
            rstCom("AU_NAME") = strAuthor
            If IsNull(strAuthorHref) = False Then
            rstCom("AU_Href").Value = strAuthorHref
            End If
            
            DATA_UpdateAuthor = rstCom("AU_ID")
            rstCom.Update
          Else
            DATA_UpdateAuthor = rstCom("AU_ID")
        End If
      Else
        '''''debug.print strSql
        DisplayInfo ("Update Author failed")
    End If
    End If
    
    On Error Resume Next
      rstCom.Close
      Set rstCom = Nothing
    On Error GoTo 0

End Function








'DATA_UPDATEFIELD 1, "PS_SERVERFILE", ""

Sub DATA_UPDATEFIELD(lngPS_ID As Long, strFieldname As String, StrValue As String)
  
  Dim iResult As Integer
  Dim rstCom As adodb.Recordset
  Dim strSql As String
  
    strSql = "SELECT PS_ID, " & strFieldname & " FROM tblPS_PSC WHERE PS_ID = " & CStr(lngPS_ID)
    
    iResult = Ado.J2_Connect(gStrDatabaseFilename, Access2000)
    If iResult = J2_ADO.EnError.No_Errors Then
        iResult = Ado.J2_Recordset(strSql, adOpenDynamic, adLockOptimistic, adCmdText, rstCom)
        If iResult = J2_ADO.EnError.No_Errors Then
            If Not (rstCom.EOF And rstCom.BOF) Then
                    rstCom(strFieldname).Value = StrValue
                    rstCom.Update
            End If
      Else
        '''''debug.print strSql
        DisplayInfo ("Update " & strFieldname & " in PSC failed (ID:" & lngPS_ID & " VALUE:" & StrValue & " ")
    End If
    End If
    
    On Error Resume Next
      rstCom.Close
      Set rstCom = Nothing
        Ado.J2_Disconnect
    On Error GoTo 0

End Sub
