Attribute VB_Name = "modStringManipulatie"

Option Explicit
Global StrPosAfterSearch As Long
Global strPosStartpos As Long

Public Function STRING_GetStringBetween(ByVal strSearchIn As String, ByVal strFrom As String, ByVal strUntil As String, _
                                        Optional ByVal lStartAtPos As Long = 0) As String

  ' This function gets in a string and two keywords
  ' and returns the string between the keywords
    
  Dim s1 As Long
  Dim s2 As Long
  Dim s  As Long
  Dim L As Long
  Dim foundstr As String

    s1 = InStr(lStartAtPos + 1, strSearchIn, strFrom, vbBinaryCompare)
    s2 = InStr(s1 + 1, strSearchIn, strUntil, vbBinaryCompare)
    
    If s1 = 0 Or s2 = 0 Or IsNull(s1) Or IsNull(s2) Then
        foundstr = "<NotFound>"
      Else
        s = s1 + Len(strFrom)
        L = s2 - s
        foundstr = Mid$(strSearchIn, s, L)
    End If
    
    STRING_GetStringBetween = foundstr
    If s + L > 0 Then
        strPosStartpos = s
        StrPosAfterSearch = (s + L) - 1
    End If

End Function

Function STRING_Uitlijnen(str As String) As String

    STRING_Uitlijnen = str & Space$(20 - (Len(str)))

End Function

Function STRING_CountArgs(strArg As String) As Integer

  Dim intTeller As Integer
  Dim intPos As Integer
  Dim blnFirst As Boolean

    '====================================================================================
    '  Versie:   1.0
    '  Auteur:   Johan Elzer
    '  Datum:    16 april 1996
    '
    '  Omschrijving: Deze functie telt het aantal argumenten die in de string 'gestopt' zijn.
    '                Bij voorkeur zo aanpassen dat het ook goed werkt als de slot-puntkomma vergeten is
    '====================================================================================
   
    intPos = 0
    intTeller = 0
    blnFirst = True
   
    'controle op ;
    If Not (Right$(strArg, 1) = ";") Then
        strArg = strArg & ";"
    End If
   
    Do While intPos > 0 Or blnFirst
        blnFirst = False
        intPos = InStr(intPos + 1, strArg, ";")
        If intPos > 0 Then
            intTeller = intTeller + 1
        End If
    Loop
   
    STRING_CountArgs = intTeller
   
End Function

Function STRING_GeefArg(ByVal intArg As Integer, ByVal strCompl As String) As String

  Dim intStart As Integer
  Dim intTeller As Integer
  Dim strDummy As String

    '====================================================================================
    '  Versie:   1.0
    '  Auteur:   E.R. Toonen, Rijnhaave Office Automation
    '  Datum:    21 Augustus 1996
    '
    '  Omschrijving: De functie haalt uit een string van argumenten het gewenste argument.
    '                Het scheidingsteken tussen de argumenten is ; (punt-komma)
    '====================================================================================
   
    intStart = 1
    intTeller = 0
    '-------------------------------------------
    'Test of de argumentenlijst eindigt op een ;
    '-------------------------------------------
    If Right$(strCompl, 1) <> ";" Then
        strCompl = strCompl & ";"
    End If

    '------------------------------------
    'Pak het juiste argument uit de lijst
    '------------------------------------
    Do Until intTeller = intArg - 1
        intStart = InStr(intStart, strCompl, ";")
        If intStart = 0 Then
            '---------------------------------------------------------------------
            'Het aantal meegegeven argumenten is kleiner dan het gewenste argument
            '---------------------------------------------------------------------
            GoTo STRING_GeefArg_Err
          Else
            '-------------------------
            'Start voorbij de ; zetten
            '-------------------------
            intStart = intStart + 1
        End If
        intTeller = intTeller + 1
    Loop
    If (InStr(intStart, strCompl, ";") - intStart) < 1 Then
        '------------------------------------
        'volgende argument niet meer gevonden
        '------------------------------------
        strDummy = ""
      Else
        strDummy = Trim$(Mid$(strCompl, intStart, InStr(intStart, strCompl, ";") - intStart))
    End If
   
    If strDummy = "" Then
        STRING_GeefArg = ""
      Else
        STRING_GeefArg = strDummy
    End If
   
STRING_GeefArg_Exit:

Exit Function
   
STRING_GeefArg_Err:
    STRING_GeefArg = ""
    Resume STRING_GeefArg_Exit
   
End Function

'**************************************
' Name: A Split Procedure
' Description:Splits a string into an ar
'     ray. If you send a " " it will split all
'     the words into each array position.
' By: Paul Spiteri
'
' Inputs:The string to split.
'The STRING_ArraySplit, e.g. " "
'
' Returns:aantal opgeslagen velden
'     s.
'
' Assumes:Private Sub Command1_Click()
'Sub testa()
'Dim SplitReturn As Variant
''''''debug.print STRING_ArraySplit("t,ssst,t,t,t,e,e,e,e,", ",", SplitReturn)
'MsgBox SplitReturn(2)
'End Sub

'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/xq/ASP/txtCode
'     Id.9165/lngWId.1/qx/vb/scripts/ShowCode.
'     htm'for details.'**************************************

Public Function STRING_ArraySplit(SplitString As String, SplitLetter As String, ByRef SplitVariant As Variant) As Long
    
  Dim TempLetter As String
  Dim TempSplit As String
  Dim i As Integer
  Dim x As Integer
  Dim StartPos As Integer

    ReDim SplitArray(1 To 1) As Variant
    SplitString = SplitString & SplitLetter

    For i = 1 To Len(SplitString)
        TempLetter = Mid$(SplitString, i, Len(SplitLetter))

        If TempLetter = SplitLetter Then
            TempSplit = Mid$(SplitString, (StartPos + 1), (i - StartPos) - 1)

            If TempSplit <> "" Then
                x = x + 1
                ReDim Preserve SplitArray(1 To x) As Variant
                SplitArray(x) = TempSplit
            End If
            StartPos = i
        End If
    Next i
    STRING_ArraySplit = x
    SplitVariant = SplitArray
    
End Function


Public Function GetFilePath(ByVal sFilename As String, Optional ByVal bAddBackslash As Boolean) As String

  'Returns Path Without FileTitle
  
  Dim lPos As Long

    lPos = InStrRev(sFilename, "/")

    If lPos > 0 Then
        GetFilePath = Left$(sFilename, lPos - 1) _
                      & IIf(bAddBackslash, "/", "")
      Else
        GetFilePath = ""
    End If
    
End Function



Public Function StringStripper(strIn) As String
Dim strCopy As String

Dim strResult As String
Dim intcharpos As Integer
Dim intAsc As Integer

strCopy = strIn


For intAsc = 0 To 255
If intAsc = 32 Then intAsc = 33         ' <space>
If intAsc = 48 Then intAsc = 58         ' 0 - 9
If intAsc = 65 Then intAsc = 91         ' A - Z
If intAsc = 97 Then intAsc = 123        ' a - z

''debug.print intAsc, Chr$(intAsc)

Do While InStr(1, strCopy, Chr$(intAsc)) > 0
    intcharpos = InStr(1, strCopy, Chr$(intAsc))
    
    If intcharpos > 0 Then
        strResult = Mid$(strCopy, 1, intcharpos - 1)
        strCopy = strResult & Mid$(strCopy, intcharpos + 1)
    End If
Loop


Next intAsc





'strResult = Replace(strIn, Chr$(122), " ")


StringStripper = strCopy
''debug.print Len(strIn), Len(strCopy)



End Function

