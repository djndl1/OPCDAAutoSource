VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ItemValuesDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Values"
   ClientHeight    =   4545
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView ItemsLV 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6800
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton CloseBTN 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5175
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "ItemValuesDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============================================================================
' TITLE: ItemValuesDlg.frm
'
' CONTENTS:
'
' A dialog used to display a set of item values.
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
'closes the form.
Private Sub CloseBTN_Click()
    Me.Hide
End Sub
'shows a list of item values.
Public Sub ShowItems(ByRef owner As Form, ByVal count As Long, ByRef items() As OPCItem, ByRef values() As Variant, ByRef Qualities() As Long, ByRef TimeStamps() As Date, ByRef Errors() As Long)

    'clear existing contents
    ItemsLV.ListItems.Clear
    ItemsLV.ColumnHeaders.Clear
    
    'add list columns.
    ItemsLV.ColumnHeaders.Add , , "Item ID", 2500
    ItemsLV.ColumnHeaders.Add , , "Value", 2500
    ItemsLV.ColumnHeaders.Add , , "Quality", 1000
    ItemsLV.ColumnHeaders.Add , , "Timestamp", 2000
    ItemsLV.ColumnHeaders.Add , , "Error", 1000
       
    'add list items.
    Dim ii As Integer
    Dim item As OPCItem
    Dim listItem As listItem
    
    For ii = 1 To count
    
        Set listItem = ItemsLV.ListItems.Add
        
        listItem.Text = items(ii).ItemID
                
        If VarType(values(ii)) > vbArray Then
            listItem.SubItems(1) = "Array(" + CStr(UBound(values(ii)) - 1) + ")"
        Else
            listItem.SubItems(1) = values(ii)
        End If
        
        listItem.SubItems(2) = CStr(Qualities(ii))
        listItem.SubItems(3) = TimeStamps(ii)
        listItem.SubItems(4) = CStr(Errors(ii))
    
    Next ii

    Show vbModal, owner
    
End Sub
'shows a list of item errors.
Public Sub ShowErrors(ByRef owner As Form, ByVal count As Long, ByRef items() As OPCItem, ByRef values() As Variant, ByRef Errors() As Long)

    'clear existing contents
    ItemsLV.ListItems.Clear
    ItemsLV.ColumnHeaders.Clear
    
    'add list columns.
    ItemsLV.ColumnHeaders.Add , , "Item ID", 2500
    ItemsLV.ColumnHeaders.Add , , "Value", 2500
    ItemsLV.ColumnHeaders.Add , , "Error", 1000
       
    'add list items.
    Dim ii As Integer
    Dim item As OPCItem
    Dim listItem As listItem
    
    For ii = 1 To count
    
        If Not Errors(ii) = 0 Then
        
            Set listItem = ItemsLV.ListItems.Add
            
            listItem.Text = items(ii).ItemID
                    
            If VarType(values(ii)) > vbArray Then
                listItem.SubItems(1) = "Array(" + CStr(UBound(values(ii)) - 1) + ")"
            Else
                listItem.SubItems(1) = values(ii)
            End If
            
            listItem.SubItems(2) = CStr(Errors(ii))
        
        End If
    
    Next ii

    'only show dialog if errors exist.
    If ItemsLV.ListItems.count > 0 Then
        Show vbModal, owner
    Else
        MsgBox "No errors found.", vbOKOnly, "Show Item Errors"
    End If
    
End Sub

