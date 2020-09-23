Attribute VB_Name = "modDisplayinfo"
Option Explicit
Global m_lngTmrCleanUptext As Integer     'Clean up lblstatusbar after 10 sec
Global m_lngTmrCleanUptext1 As Integer     'Clean up lblstatusbar1 after 10 sec

Sub DisplayInfo(strText As String)
    Form1.lblDisplayInfo.Caption = strText
End Sub

Sub DisplayInfoNOcrlf(strText As String)
Form1.lblDisplayInfo.Caption = Form1.lblDisplayInfo.Caption & strText
    If g_bFeedback = 1 Then
        
    End If

End Sub

Sub DisplayInfoLbl(strText As String)
 Form1.lblDisplayInfo.Caption = strText
End Sub

Sub DisplayInfoLbl1(strText As String)
    m_lngTmrCleanUptext1 = 0
   DoEvents
End Sub

Sub DisplayInfoCls()

    If g_bFeedback = 1 Then
    End If

End Sub

Sub DisplayInfoProgressbar1(lngValue As Long, lngMax As Long)
    If lngMax > 0 Then
     On Error Resume Next
   Form1.ctrlProgressBar1.SetMeter (100 / lngMax * lngValue)
   End If
   
   Form1.ctrlProgressBar1.Visible = True

   
   
End Sub

Sub DisplayInfoProgressbar1Hide()
   Form1.ctrlProgressBar1.Visible = False
   DoEvents
End Sub


