VERSION 5.00
Begin VB.Form SelectServerDlg 
   Caption         =   "Select Server"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton OkBTN 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton CancelBTN 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.ListBox ServersLB 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "SelectServerDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============================================================================
' TITLE: SelectServerDlg.frm
'
' CONTENTS:
'
' A dialog used to select a server to connect to.
'
' (c) Copyright 2003-2004 The OPC Foundation
' ALL RIGHTS RESERVED.
'
' DISCLAIMER:
'  This code is provided by the OPC Foundation solely to assist in
'  understanding and use of the appropriate OPC Specification(s) and may be
'  used as set forth in the License Grant section of the OPC Specification.
'  This code is provided as-is and without warranty or support of any sort
'  and is subject to the Warranty and Liability Disclaimers which appear
'  in the printed OPC Specification.
'
' MODIFICATION LOG:
'
' Date       By    Notes
' ---------- ---   -----
' 2003/12/01 RSA   Initial implementation.

Option Explicit
Option Base 1
'shows the server status
Public Function SelectServer(ByRef owner As Form, ByRef node As String) As String

On Error GoTo Failed
      
    'clear the list box.
    ServersLB.Clear
    
    'create a new server object.
    Dim server As OPCServer
    Set server = New OPCServer
        
    'fetch the prog ids of the known servers.
    Dim servers() As String
    servers = server.GetOPCServers(node)
    
    'update the list of servers.
    Dim ii As Long
        
    For ii = 1 To UBound(servers)
        ServersLB.AddItem servers(ii)
    Next ii
        
    'show dialog.
    Me.Show vbModal, owner
    
    'find selected server.
    SelectServer = ""
    
    For ii = 1 To UBound(servers)
        If ServersLB.Selected(ii - 1) = True Then
            SelectServer = servers(ii)
            Exit For
        End If
    Next ii
        
    Exit Function

Failed:

    MsgBox "Could not select server.\r\n" + Err.Description

End Function
Private Sub OkBTN_Click()
    Me.Hide
End Sub
Private Sub CancelBTN_Click()
    Me.Hide
End Sub
Private Sub ServersLB_DblClick()
    Me.Hide
End Sub
