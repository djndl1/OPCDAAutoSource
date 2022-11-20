VERSION 5.00
Begin VB.Form ViewStatusDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Server Status"
   ClientHeight    =   4200
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.Timer RefreshTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   120
      Top             =   3720
   End
   Begin VB.TextBox ServerVersionTB 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1395
      Width           =   4215
   End
   Begin VB.TextBox StartTimeTB 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1710
      Width           =   4215
   End
   Begin VB.TextBox CurrentTimeTB 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2040
      Width           =   4215
   End
   Begin VB.TextBox HostNameTB 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   120
      Width           =   4215
   End
   Begin VB.TextBox ServerNameTB 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   435
      Width           =   4215
   End
   Begin VB.TextBox VendorInfoTB 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   765
      Width           =   4215
   End
   Begin VB.TextBox ServerStateTB 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox LastUpdateTimeTB 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2355
      Width           =   4215
   End
   Begin VB.TextBox BandwidthTB 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2670
      Width           =   4215
   End
   Begin VB.TextBox LocaleTB 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2985
      Width           =   4215
   End
   Begin VB.TextBox ClientNameTB 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3315
      Width           =   4215
   End
   Begin VB.CommandButton CloseBTN 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   2348
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label ServerLBs 
      AutoSize        =   -1  'True
      Caption         =   "Server Name"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   480
      Width           =   930
   End
   Begin VB.Label ServerLBs 
      AutoSize        =   -1  'True
      Caption         =   "Vendor Info"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   810
      Width           =   825
   End
   Begin VB.Label ServerLBs 
      AutoSize        =   -1  'True
      Caption         =   "Host Name"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   165
      Width           =   795
   End
   Begin VB.Label ServerLBs 
      AutoSize        =   -1  'True
      Caption         =   "Server State"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   1125
      Width           =   885
   End
   Begin VB.Label ServerLBs 
      AutoSize        =   -1  'True
      Caption         =   "Server Version"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label ServerLBs 
      AutoSize        =   -1  'True
      Caption         =   "Start Time"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   1755
      Width           =   720
   End
   Begin VB.Label ServerLBs 
      AutoSize        =   -1  'True
      Caption         =   "Current Time"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   2085
      Width           =   900
   End
   Begin VB.Label ServerLBs 
      AutoSize        =   -1  'True
      Caption         =   "Last Update Time"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Label ServerLBs 
      AutoSize        =   -1  'True
      Caption         =   "Bandwidth"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   14
      Top             =   2715
      Width           =   750
   End
   Begin VB.Label ServerLBs 
      AutoSize        =   -1  'True
      Caption         =   "Locale"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   13
      Top             =   3030
      Width           =   480
   End
   Begin VB.Label ServerLBs 
      AutoSize        =   -1  'True
      Caption         =   "Client Name"
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   855
   End
End
Attribute VB_Name = "ViewStatusDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============================================================================
' TITLE: ViewStatusDlg.frm
'
' CONTENTS:
'
' A dialog used to view the current server status.
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
Dim theServer As OPCServer
'closes the dialog
Private Sub CloseBTN_Click()
    RefreshTimer.Enabled = False
    Me.Hide
End Sub
'updates controls
Private Sub UpdateStatus(ByRef server As OPCServer)
    
On Error GoTo Failed

    HostNameTB.Text = server.ServerNode
    
    If HostNameTB.Text = "" Then
        HostNameTB.Text = "localhost"
    End If
    
    ServerNameTB.Text = server.ServerName
    VendorInfoTB.Text = server.VendorInfo
    ServerVersionTB.Text = CStr(server.MajorVersion) + "." + CStr(server.MinorVersion) + "." + CStr(server.BuildNumber)
    ServerStateTB.Text = CStr(server.ServerState)
    StartTimeTB.Text = server.StartTime
    CurrentTimeTB.Text = server.CurrentTime
    LastUpdateTimeTB.Text = server.LastUpdateTime
    
    BandwidthTB.Text = ""
    
    If Not server.Bandwidth = -1 Then
        BandwidthTB.Text = CStr(server.Bandwidth)
    End If
    
    LocaleTB.Text = CStr(server.LocaleID)
    ClientNameTB.Text = server.ClientName
    
    Exit Sub
    
Failed:
    
    RefreshTimer.Enabled = False
    
End Sub
'shows the server status
Public Sub ViewStatus(ByRef owner As Form, ByRef server As OPCServer)

On Error GoTo Failed
    
    'initial  status
    Set theServer = server
    UpdateStatus server
    
    'start refresh timer.
    RefreshTimer.Enabled = True
    
    'show dialog.
    Me.Show vbModal, owner
    
    Exit Sub

Failed:

    MsgBox "Could not get server status."

End Sub
'updates the status when the timer expires.
Private Sub RefreshTimer_Timer()
    UpdateStatus theServer
End Sub
