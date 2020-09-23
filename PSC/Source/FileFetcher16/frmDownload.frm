VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmDownload 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TmrInet 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1260
      Top             =   780
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   330
      Top             =   765
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   300
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Integer
Private Sub Form_Load()

    Me.Caption = g_ApplTitle

End Sub


Private Sub Inet_StateChanged(ByVal State As Integer)
  Dim strInfotext As String

    Select Case State
      Case 0 'icNone
        strInfotext = "No state to report."
      Case 1 'icHostResolvingHost
        strInfotext = "The control is looking up the IP address of the specified host computer."
      Case 2 'icHostResolved
        strInfotext = "The control successfully found the IP address of the specified host computer."
      Case 3 'icConnecting
        strInfotext = "The control is connecting to the host computer."
      Case 4 'icConnected
        strInfotext = "The control successfully connected to the host computer."
      Case 5 'icRequesting
        strInfotext = "The control is sending a request to the host computer."
      Case 6 'icRequestSent
        strInfotext = "The control successfully sent the request."
      Case 7 'icReceivingResponse
        strInfotext = "The control is receiving a response from the host computer."
      Case 8 'icResponseReceived
        strInfotext = "The control successfully received a response from the host computer."
      Case 9 'icDisconnecting
        strInfotext = "The control is disconnecting from the host computer."
      Case 10 'icDisconnected
        strInfotext = "The control successfully disconnected from the host computer."
      Case 11 'icError
        strInfotext = "An error occurred in communicating with the host computer."
      Case 12 'icResponseCompleted
        strInfotext = "The request has completed and all data has been received."
      Case Else
        strInfotext = "Unknown state"
    End Select
    DisplayInfo strInfotext

End Sub

Private Sub TmrInet_Timer()
x = x + 1



    If Me.Inet.StillExecuting Then
    Select Case x
    Case 1
   '     DisplayInfo ("|")
    Case 2
  '      DisplayInfo ("/")
    Case 3
 '       DisplayInfo ("-")
        
    Case 4
'        DisplayInfo ("\")
    End Select
    
      '  DisplayInfoNOcrlf (".")
     
     End If
If x = 5 Then x = 0

End Sub





Public Function GetDownloadUrl(lngPS_ID As Long, strUrl As String, strTitle As String) As String
'Return a .zip path


  Dim strTemp As String
  Dim strFilterZip As String
  Dim strAuthor As String
  Dim strAuthorHref As String
    If Inet.StillExecuting Then
        DisplayInfo "The internet control is still executing"
        GetDownloadUrl = ""
        Exit Function
    End If
    
    TmrInet.Enabled = True
    DisplayInfo ("The internet control is busy")
    If Len(strUrl) < 5 Then
    MsgBox "No URL found"
    Exit Function
    End If
    On Error Resume Next
    strTemp = Inet.OpenURL(strUrl)
 
    DisplayInfo ("Filtering URL from :" & strUrl)

    'Get filename of downloadable zip file
    strFilterZip = STRING_GetStringBetween(strTemp, "<a name=""zip"">", "<img border=""0""")
    strFilterZip = STRING_GetStringBetween(strFilterZip, """", """")
    
    'get author-name out of the page
    strAuthor = Trim(STRING_GetStringBetween(strTemp, "<b>By:</b>", "<br>"))
    If Trim(STRING_GetStringBetween(strAuthor, "<a href=""", """>")) <> "<NotFound>" Then
        strAuthorHref = Trim(STRING_GetStringBetween(strAuthor, "<a href=""", """>"))
        strAuthor = STRING_GetStringBetween(strAuthor, """>", "</a>")
        strAuthor = Replace(strAuthor, Chr$(9), " ")
        strAuthor = Replace(strAuthor, Chr$(10), " ")
        strAuthor = Replace(strAuthor, Chr$(13), " ")
        strAuthor = Trim(strAuthor)
    End If
         
    DATA_UPDATEFIELD lngPS_ID, "PS_AU_ID", DATA_UpdateAuthor(strAuthor, strAuthorHref)
   
    'did we found a zip
    If strFilterZip <> "<NotFound>" Then
    
        DATA_UPDATEFIELD lngPS_ID, "PS_SERVERFILE", strFilterZip
        DATA_UPDATEFIELD lngPS_ID, "PS_LOCALDIR", File.RemoveBackSlash(gStrDefaultDownloadDir) & "\" & StringStripper(strTitle) & ".zip"
        DownloadFile strFilterZip, File.RemoveBackSlash(gStrDefaultDownloadDir) & "\" & StringStripper(strTitle) & ".zip"
        DATA_UPDATEFIELD lngPS_ID, "PS_DOWNLOADED", "Y"
    Else
        'No zip found create text file with copy and paste friendly code :)
        strFilterZip = STRING_GetStringBetween(strTemp, "Click here for a <a href=""", """>copy-and-paste friendly")
         
        If InStr(1, strTemp, "The author of this code has deleted it or it has been removed.") <> 0 Then
           'Code = deleted from server
           DisplayInfoLbl "The author of this code has deleted it or it has been removed."
           DATA_UPDATEFIELD lngPS_ID, "PS_DOWNLOADED", "D"
           Call File.HORAEST_AddLineToFile(g_Logfile, "[PSID:" & CStr(lngPS_ID) & "The author of this code has deleted it or it has been removed.]")
        
        Else
            'Code is txt
            If strFilterZip = "<NotFound>" Then
                Call File.HORAEST_AddLineToFile(g_Logfile, "[PSID:" & CStr(lngPS_ID) & " not found ? article ?]")
                DATA_UPDATEFIELD lngPS_ID, "PS_DOWNLOADED", "U"
            
            Else
                DATA_UPDATEFIELD lngPS_ID, "PS_SERVERFILE", "http://www.planet-source-code.com" & strFilterZip
                DATA_UPDATEFIELD lngPS_ID, "PS_LOCALDIR", File.RemoveBackSlash(gStrDefaultDownloadDir) & "\" & StringStripper(strTitle) & ".txt"
                DownloadFile "http://www.planet-source-code.com" & strFilterZip, File.RemoveBackSlash(gStrDefaultDownloadDir) & "\" & StringStripper(strTitle) & ".txt"
            End If
        End If
    
    End If
    
    DisplayInfoLbl strFilterZip
    
    TmrInet.Enabled = False

    GetDownloadUrl = strFilterZip
    

   
End Function

Public Function DownloadFile(strFromUrl As String, strToFile As String) As Long
If Me.Inet.StillExecuting = False Then
  Dim bytes() As Byte
  Dim fnum As Integer

''    Call DisplayIconInit(AnimationType.Download)
    DisplayInfoLbl "Download:" & strFromUrl
    ' Get the file.
    Me.TmrInet.Enabled = True
    bytes() = Me.Inet.OpenURL(strFromUrl, icByteArray)
    
    Me.TmrInet.Enabled = False
    ' Save the file.
    DisplayInfoLbl1 "Save:" & strToFile
    

    fnum = FreeFile
    Open strToFile For Binary Access Write As #fnum
    Put #fnum, , bytes()
    Close #fnum
    
    DisplayInfoLbl "Done..."
    DisplayInfoLbl1 " "
    
Else
    MsgBox "Still executing last request please wait"
End If

End Function








