VERSION 5.00
Begin VB.Form EditGroupDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Group"
   ClientHeight    =   3450
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox GroupNameTB 
      Height          =   285
      Left            =   1290
      TabIndex        =   10
      Top             =   120
      Width           =   3735
   End
   Begin VB.TextBox ClientHandleTB 
      Height          =   285
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   420
      Width           =   3735
   End
   Begin VB.TextBox ServerHandleTB 
      Height          =   285
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   720
      Width           =   3735
   End
   Begin VB.CheckBox IsActiveCB 
      Height          =   285
      Left            =   1290
      TabIndex        =   7
      Top             =   1020
      Width           =   855
   End
   Begin VB.CheckBox IsSubscribedCB 
      Height          =   285
      Left            =   1290
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox LocaleTB 
      Height          =   285
      Left            =   1290
      TabIndex        =   5
      Top             =   1620
      Width           =   3735
   End
   Begin VB.TextBox TimeBiasTB 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1290
      TabIndex        =   4
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox DeadbandTB 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1290
      TabIndex        =   3
      Top             =   2220
      Width           =   3735
   End
   Begin VB.TextBox UpdateRateTB 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1290
      TabIndex        =   2
      Top             =   2520
      Width           =   3735
   End
   Begin VB.CommandButton OkBTN 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton CancelBTN 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label GroupLBs 
      AutoSize        =   -1  'True
      Caption         =   "Update Rate"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   19
      Top             =   2565
      Width           =   915
   End
   Begin VB.Label GroupLBs 
      AutoSize        =   -1  'True
      Caption         =   "Deadband"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   2265
      Width           =   750
   End
   Begin VB.Label GroupLBs 
      AutoSize        =   -1  'True
      Caption         =   "Client Handle"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   465
      Width           =   945
   End
   Begin VB.Label GroupLBs 
      AutoSize        =   -1  'True
      Caption         =   "Server Handle"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   765
      Width           =   1020
   End
   Begin VB.Label GroupLBs 
      AutoSize        =   -1  'True
      Caption         =   "Active"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   1065
      Width           =   450
   End
   Begin VB.Label GroupLBs 
      AutoSize        =   -1  'True
      Caption         =   "Subscribed"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1365
      Width           =   795
   End
   Begin VB.Label GroupLBs 
      AutoSize        =   -1  'True
      Caption         =   "Locale"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   1665
      Width           =   480
   End
   Begin VB.Label GroupLBs 
      AutoSize        =   -1  'True
      Caption         =   "Time Bias"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1965
      Width           =   690
   End
   Begin VB.Label GroupLBs 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   165
      Width           =   420
   End
End
Attribute VB_Name = "EditGroupDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============================================================================
' TITLE: EditGroupDlg.frm
'
' CONTENTS:
'
' A dialog used to edit the state of a group.
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

Dim group As OPCGroup
Dim groupName As String
Dim changed As Boolean
'closes the dialog
Private Sub OkBTN_Click()

On Error GoTo Failed

    'only update if name is actually changed.
    If Not groupName = GroupNameTB.Text Then
        group.Name = GroupNameTB.Text
    End If
        
    If IsActiveCB.value = vbChecked Then
        group.IsActive = True
    Else
        group.IsActive = False
    End If
    
    If IsSubscribedCB.value = vbChecked Then
        group.IsSubscribed = True
    Else
        group.IsSubscribed = False
    End If
       
    group.LocaleID = CLng(LocaleTB.Text)
    group.TimeBias = CLng(TimeBiasTB.Text)
    group.DeadBand = CSng(DeadbandTB.Text)
    group.UpdateRate = CLng(UpdateRateTB.Text)
    
    changed = True
    Me.Hide
       
    Exit Sub
    
Failed:

    MsgBox Error
    
End Sub
'closes the dialog
Private Sub CancelBTN_Click()
    changed = False
    Me.Hide
End Sub
'shows the server status
Public Function EditGroup(ByRef owner As Form, ByRef aGroup As OPCGroup) As Boolean

On Error GoTo Failed
    
    EditGroup = False
    
    Set group = aGroup
    
    'update controls.
    GroupNameTB.Text = group.Name
    ClientHandleTB.Text = CStr(group.ClientHandle)
    ServerHandleTB.Text = CStr(group.ServerHandle)
    
    If group.IsActive Then
        IsActiveCB.value = vbChecked
    Else
        IsActiveCB.value = vbUnchecked
    End If
    
    If group.IsSubscribed Then
        IsSubscribedCB.value = vbChecked
    Else
        IsSubscribedCB.value = vbUnchecked
    End If
    
    LocaleTB.Text = CStr(group.LocaleID)
    TimeBiasTB.Text = CStr(group.TimeBias)
    DeadbandTB.Text = CStr(group.DeadBand)
    UpdateRateTB.Text = CStr(group.UpdateRate)
        
    'cache existing group name.
    groupName = GroupNameTB.Text
    
    'show dialog.
    Me.Show vbModal, owner
    EditGroup = changed
    Exit Function

Failed:

    MsgBox Error

End Function
