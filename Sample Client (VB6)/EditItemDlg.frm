VERSION 5.00
Begin VB.Form EditItemDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Item"
   ClientHeight    =   3135
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox NativeDataTypeTB 
      ForeColor       =   &H80000011&
      Height          =   285
      Left            =   1770
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox ReqDataTypeTB 
      Height          =   285
      Left            =   1770
      TabIndex        =   8
      Top             =   990
      Width           =   3735
   End
   Begin VB.TextBox AccessRightsTB 
      ForeColor       =   &H80000011&
      Height          =   285
      Left            =   1770
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1860
      Width           =   3735
   End
   Begin VB.TextBox AccessPathTB 
      ForeColor       =   &H80000011&
      Height          =   285
      Left            =   1770
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   405
      Width           =   3735
   End
   Begin VB.CheckBox IsActiveCB 
      Height          =   285
      Left            =   1770
      TabIndex        =   5
      Top             =   705
      Width           =   855
   End
   Begin VB.TextBox ServerHandleTB 
      ForeColor       =   &H80000011&
      Height          =   285
      Left            =   1770
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1575
      Width           =   3735
   End
   Begin VB.TextBox ClientHandleTB 
      ForeColor       =   &H80000011&
      Height          =   285
      Left            =   1770
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1290
      Width           =   3735
   End
   Begin VB.TextBox ItemIDTB 
      ForeColor       =   &H80000011&
      Height          =   285
      Left            =   1770
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton CancelBTN 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton OkBTN 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      Caption         =   "Item ID"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   17
      Top             =   165
      Width           =   510
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      Caption         =   "Access Rights"
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   16
      Top             =   1905
      Width           =   1020
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      Caption         =   "Access Path"
      Height          =   195
      Index           =   11
      Left            =   120
      TabIndex        =   15
      Top             =   450
      Width           =   900
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      Caption         =   "Active"
      Height          =   195
      Index           =   13
      Left            =   120
      TabIndex        =   14
      Top             =   750
      Width           =   450
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      Caption         =   "Server Handle"
      Height          =   195
      Index           =   14
      Left            =   120
      TabIndex        =   13
      Top             =   1620
      Width           =   1020
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      Caption         =   "Client Handle"
      Height          =   195
      Index           =   15
      Left            =   120
      TabIndex        =   12
      Top             =   1335
      Width           =   945
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      Caption         =   "Requested Data Type"
      Height          =   195
      Index           =   16
      Left            =   120
      TabIndex        =   11
      Top             =   1035
      Width           =   1575
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      Caption         =   "Native Data Type"
      Height          =   195
      Index           =   17
      Left            =   120
      TabIndex        =   10
      Top             =   2205
      Width           =   1260
   End
End
Attribute VB_Name = "EditItemDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============================================================================
' TITLE: EditItemDlg.frm
'
' CONTENTS:
'
' A dialog used to edit the state of an item.
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

Dim item As OPCItem
Dim changed As Boolean
'closes the dialog
Private Sub OkBTN_Click()

On Error GoTo Failed
        
    If IsActiveCB.value = vbChecked Then
        item.IsActive = True
    Else
        item.IsActive = False
    End If
    
    item.RequestedDataType = CInt(ReqDataTypeTB.Text)
    
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
'show the current item properties.
Public Function EditItem(ByRef owner As Form, ByRef anItem As OPCItem) As Boolean

On Error GoTo Failed
    
    EditItem = False
    
    Set item = anItem
    
    'update controls.
    ItemIDTB.Text = item.ItemID
    AccessPathTB.Text = item.AccessPath

    If item.IsActive Then
        IsActiveCB.value = vbChecked
    Else
        IsActiveCB.value = vbUnchecked
    End If
    
    ReqDataTypeTB.Text = CStr(item.RequestedDataType)

    ClientHandleTB.Text = CStr(item.ClientHandle)
    ServerHandleTB.Text = CStr(item.ServerHandle)
    AccessRightsTB.Text = CStr(item.AccessRights)
    NativeDataTypeTB.Text = CStr(item.CanonicalDataType)
    
    'show dialog.
    Me.Show vbModal, owner
    EditItem = changed
    Exit Function

Failed:

    MsgBox Error

End Function

