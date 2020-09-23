VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCLose 
      Caption         =   "Close"
      Height          =   510
      Left            =   3480
      TabIndex        =   0
      Top             =   2535
      Width           =   975
   End
   Begin MSMAPI.MAPISession MAPISession 
      Left            =   600
      Top             =   -15
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   -1  'True
   End
   Begin MSMAPI.MAPIMessages MAPIMessages 
      Left            =   0
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   -1  'True
      FetchUnreadOnly =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Unload(Cancel As Integer)

    Call Uitkijk_logoff
    If Me.MAPISession.SessionID > 0 Then
        Me.MAPISession.SignOff
    End If

End Sub



