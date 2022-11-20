VERSION 5.00
Begin VB.Form EditValueDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Value"
   ClientHeight    =   1095
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OkBTN 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton CancelBTN 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox ValueTB 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      Caption         =   "Value"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   3
      Top             =   165
      Width           =   405
   End
End
Attribute VB_Name = "EditValueDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============================================================================
' TITLE: EditValueDlg.frm
'
' CONTENTS:
'
' A dialog used to edit the value of an item.
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

Dim value As String
Dim datatype As VbVarType
Dim changed As Boolean
'closes the dialog
Private Sub OkBTN_Click()

On Error GoTo Failed:
     
    Select Case datatype
        Case vbBoolean
            value = CBool(ValueTB.Text)
        Case vbByte
            value = CByte(ValueTB.Text)
        Case vbCurrency
            value = CCur(ValueTB.Text)
        Case vbDate
            value = CDate(ValueTB.Text)
        Case vbDouble
            value = CDbl(ValueTB.Text)
        Case vbDecimal
            value = CDec(ValueTB.Text)
        Case vbInteger
            value = CInt(ValueTB.Text)
        Case vbLong
            value = CLng(ValueTB.Text)
        Case vbSingle
            value = CSng(ValueTB.Text)
        Case vbString
            value = CStr(ValueTB.Text)
        Case vbVariant
            value = CVar(ValueTB.Text)
    End Select
    
    changed = True
    Me.Hide
    
    Exit Sub

Failed:

    MsgBox "Could not convert value to original data type."
    
End Sub
'closes the dialog
Private Sub CancelBTN_Click()
    value = ""
    changed = False
    Me.Hide
End Sub
'prompts the user to edit the value.
Public Function EditValue(ByRef owner As Form, ByVal currentValue As Variant) As Variant
    
On Error GoTo Failed:
  
    EditValue = Empty
        
    datatype = VarType(currentValue)
    ValueTB.Text = CStr(currentValue)
     
    Me.Show vbModal, owner
    
    If changed Then
        EditValue = value
    End If
    
    Exit Function
    
Failed:

    MsgBox "Could not convert value to string."
    
End Function

