Attribute VB_Name = "modFolderBrowse"
'This module contains all the declarations to use the
'Windows 95 Shell API to use the browse for folders
'dialog box.  To use the browse for folders dialog box,
'please call the BrowseForFolders function using the
'syntax: stringFolderPath=BrowseForFolders(Hwnd,TitleOfDialog)
'


Option Explicit

Public Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Const SW_SHOW = 1

Public Const BIF_RETURNONLYFSDIRS = 1 ':( As Integer ?
Public Const MAX_PATH = 260 ':( As Integer ?

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function BrowseForFolder(hwndOwner As Long, sPrompt As String) As String
     
  'declare variables to be used
  
  Dim iNull As Integer
  Dim lpIDList As Long
  Dim lResult As Long
  Dim sPath As String
  Dim udtBI As BrowseInfo

    'initialise variables
    With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Call the browse for folder API
    lpIDList = SHBrowseForFolder(udtBI)
     
    'get the resulting string path
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1) ':( Expand Structure
    End If

    'If cancel was pressed, sPath = ""
    BrowseForFolder = sPath

End Function




' Invoke the application associated with a
' document type.
Public Sub Navigate(frm As Form, ByVal document_name As String)

    ShellExecute frm.hwnd, "open", document_name, "", "", SW_SHOW

End Sub


