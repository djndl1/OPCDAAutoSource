VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BrowseItemsDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browse Items"
   ClientHeight    =   4455
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   12390
   ShowInTaskbar   =   0   'False
   Begin VB.Frame SelectedItemsFM 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   4560
      TabIndex        =   14
      Top             =   480
      Width           =   7575
      Begin VB.CommandButton RemoveBTN 
         Caption         =   "<<"
         Height          =   375
         Left            =   0
         TabIndex        =   18
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton AddBTN 
         Caption         =   ">>"
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   840
         Width           =   615
      End
      Begin VB.ListBox SelectedItemsLB 
         Height          =   2985
         Left            =   720
         TabIndex        =   15
         Top             =   120
         Width           =   6735
      End
   End
   Begin VB.CommandButton CancelBTN 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   11040
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton OkBTN 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin MSComctlLib.TreeView BrowserTV 
      Height          =   3735
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6588
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Frame PropertiesFM 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   4560
      TabIndex        =   2
      Top             =   480
      Width           =   7575
      Begin MSComctlLib.ListView PropertiesLV 
         Height          =   2775
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4895
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
      Begin VB.TextBox ItemIDTB 
         Height          =   285
         Left            =   600
         TabIndex        =   3
         Top             =   75
         Width           =   6855
      End
      Begin VB.Label ItemIDLB 
         AutoSize        =   -1  'True
         Caption         =   "Item ID"
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   510
      End
   End
   Begin MSComctlLib.TabStrip BrowseTabs 
      Height          =   3735
      Left            =   4440
      TabIndex        =   16
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6588
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame FiltersFM 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   2280
      TabIndex        =   7
      Top             =   3360
      Visible         =   0   'False
      Width           =   4335
      Begin VB.ComboBox DataTypeFilterCB 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   315
         Width           =   3135
      End
      Begin VB.ComboBox AccessRightsFilterCB 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   660
         Width           =   3135
      End
      Begin VB.TextBox NameFilterTB 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label AccessRightsFilterLB 
         AutoSize        =   -1  'True
         Caption         =   "Access Rights"
         Height          =   195
         Left            =   0
         TabIndex        =   13
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label DataTypeFilterLB 
         AutoSize        =   -1  'True
         Caption         =   "Data Type"
         Height          =   195
         Left            =   0
         TabIndex        =   12
         Top             =   375
         Width           =   750
      End
      Begin VB.Label NameFilterLB 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   0
         TabIndex        =   11
         Top             =   45
         Width           =   420
      End
   End
End
Attribute VB_Name = "BrowseItemsDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============================================================================
' TITLE: BrowseItemsDlg.frm
'
' CONTENTS:
'
' A dialog used to browse the items in a server's address space.
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

Dim server As OPCServer
Dim browser As OPCBrowser
Dim seed As Long
Dim browseOnly As Boolean
'generates a unique key.
Private Function CreateKey(ByVal prefix As String) As String
    seed = seed + 1
    CreateKey = prefix + " [" + CStr(seed) + "]"
End Function

Private Sub BrowserTV_Expand(ByVal node As MSComctlLib.node)

    'browse for children if node contains a dummy child.
    If node.Children = 1 Then
        If node.child.key = CStr(node.key) + "Dummy" Then
            BrowserTV.Nodes.Remove CStr(node.key) + "Dummy"
            ExpandBranch node
        End If
    End If
    
End Sub

'clears the selected item list
Private Sub CancelBTN_Click()
    SelectedItemsLB.Clear
    Me.Hide
End Sub
'hides the dialog
Private Sub OkBTN_Click()
    Me.Hide
End Sub
'returns the currently selected element object if any.
Private Function GetSelectedElement() As OPCBrowseElement

    Set GetSelectedElement = Nothing

    'check if anything is selected.
    If BrowserTV.SelectedItem Is Nothing Then
        Exit Function
    End If
    
    'check if a aServer object is selected.
    If Not TypeOf BrowserTV.SelectedItem.Tag Is OPCBrowseElement Then
        Exit Function
    End If
    
    'return the current selection.
    Set GetSelectedElement = BrowserTV.SelectedItem.Tag

End Function
'adds current item to selected items list.
Private Sub BrowserTV_DblClick()
    Add
End Sub
'adds current selection to item list.
Private Sub AddBTN_Click()
    Add
End Sub
'adds current selection to item list.
Private Sub Add()
    
    'check if only browsing items.
    If browseOnly Then
        Exit Sub
    End If
    
    'get currently selected element.
    Dim element As OPCBrowseElement
    Set element = GetSelectedElement()
    
    If element Is Nothing Then
        Exit Sub
    End If
    
    If Not element.ItemID = "" Then
            
        'check for duplicate item
        Dim ii As Integer
        
        For ii = 0 To SelectedItemsLB.ListCount - 1
            If SelectedItemsLB.List(ii) = element.ItemID Then
                Exit Sub
            End If
        Next ii
        
        'add item
        SelectedItemsLB.AddItem element.ItemID
        
        'show selected items tab.
        If Not BrowseTabs.SelectedItem = BrowseTabs.Tabs(2) Then
            BrowseTabs.SelectedItem = BrowseTabs.Tabs(2)
        End If
    
    End If
    
End Sub
'removes an item from the selection list.
Private Sub RemoveBTN_Click()
    Remove
End Sub
'updates the properties display
Private Sub SelectedItemdLB_DblClick()
    Remove
End Sub
'removes an item from the selection list.
Private Sub Remove()
    
    'check if only browsing items.
    If browseOnly Then
        Exit Sub
    End If
    
    If SelectedItemsLB.SelCount = 0 Then
        Exit Sub
    End If

    'remove first selected item.
    SelectedItemsLB.RemoveItem SelectedItemsLB.ListIndex
    
End Sub
'updates the properties display
Private Sub BrowserTV_Click()
    
    If server Is Nothing Or BrowserTV.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    ShowProperties BrowserTV.SelectedItem, server
    
End Sub
'initializes the form
Private Sub Form_Load()
       
    BrowseTabs.Tabs.Add
    BrowseTabs.Tabs(1).Caption = "Properties"
    BrowseTabs.Tabs(2).Caption = "Items"
    BrowseTabs.SelectedItem = BrowseTabs.Tabs(1)
       
    AccessRightsFilterCB.AddItem "readable"
    AccessRightsFilterCB.AddItem "writeable"
    AccessRightsFilterCB.AddItem "readWriteable"
    
End Sub
'handles changes to the current tab.
Private Sub BrowseTabs_Click()
    
    If BrowseTabs.SelectedItem.Index = 1 Then
        PropertiesFM.Visible = True
        SelectedItemsFM.Visible = False
    Else
        PropertiesFM.Visible = False
        SelectedItemsFM.Visible = True
    End If

End Sub
'gets the tree depth for the current node.
Private Function GetDepth(ByRef node As MSComctlLib.node) As Integer
    
    If node Is Nothing Then
        GetDepth = 0
        Exit Function
    End If
       
    GetDepth = GetDepth(node.parent) + 1
    
End Function
'gets set of branch names that point to the node.
Private Sub GetBranches(ByRef node As MSComctlLib.node, ByRef branches() As String, Index As Integer)
    
    If Not node.parent Is Nothing Then
        GetBranches node.parent, branches, Index - 1
    End If
    
    branches(Index) = node.Text

End Sub
'fetches and displays the contents of the branch.
Private Sub ExpandBranch(ByRef parent As node)
        
    Dim depth As Integer
    Dim branches() As String
    Dim ii As Integer
    
    'get current tree depth.
    depth = GetDepth(parent)
    
    If depth = 0 Then
        Exit Sub
    End If
        
    'get branch list
    ReDim branches(depth)
    GetBranches parent, branches, depth
        
    'clear existing children
    Do While parent.Children > 0
        BrowserTV.Nodes.Remove parent.child.key
    Loop
    
    'get branches from OPC aServer.
    browser.MoveTo branches
    browser.ShowBranches
            
    Dim childNode As node
    Dim childElement As OPCBrowseElement
        
    For ii = 1 To browser.count
    
        Set childElement = New OPCBrowseElement
        
        childElement.ElementName = browser.item(ii)
        childElement.ItemID = browser.GetItemID(browser.item(ii))
        childElement.HasChildren = True
        
        Set childNode = BrowserTV.Nodes.Add(parent.key, tvwChild, CreateKey(childElement.ElementName), childElement.ElementName)
        Set childNode.Tag = childElement
        
        'add a dummy child element.
        BrowserTV.Nodes.Add childNode.key, tvwChild, CStr(childNode.key) + "Dummy", ""
        
    Next ii
        
    'get leaves from OPC aServer.
    browser.ShowLeafs
    
    For ii = 1 To browser.count
        
        Set childElement = New OPCBrowseElement
        
        childElement.ElementName = browser.item(ii)
        childElement.ItemID = browser.GetItemID(browser.item(ii))
        childElement.HasChildren = False
        
        Set childNode = BrowserTV.Nodes.Add(parent.key, tvwChild, CreateKey(childElement.ElementName), childElement.ElementName)
        Set childNode.Tag = childElement
        
    Next ii

End Sub
'fetches and displays the properties for the node.
Private Sub ShowProperties(ByRef selection As node, ByRef aServer As OPCServer)
        
On Error GoTo Failed
    
    Dim ii As Long
    Dim count As Long
    Dim propertyIDs() As Long
    Dim descriptions() As String
    Dim dataTypes() As Integer
    Dim values() As Variant
    Dim Errors() As Long
            
    If selection Is Nothing Then
        Exit Sub
    End If
    
    Dim element As OPCBrowseElement
    Set element = selection.Tag
    
    If element Is Nothing Then
        Exit Sub
    End If
    
    'clear current properties.
    PropertiesLV.ListItems.Clear
    PropertiesLV.ColumnHeaders.Clear
            
    aServer.QueryAvailableProperties element.ItemID, count, propertyIDs, descriptions, dataTypes
    aServer.GetItemProperties element.ItemID, count, propertyIDs, values, Errors
    
    If count = 0 Then
        Exit Sub
    End If
    
    ItemIDTB.Text = element.ItemID
    
    'add headers
    PropertiesLV.ColumnHeaders.Add , , "ID", 500
    PropertiesLV.ColumnHeaders.Add , , "Name", 2200
    PropertiesLV.ColumnHeaders.Add , , "Type", 600
    PropertiesLV.ColumnHeaders.Add , , "Value", 2200
    PropertiesLV.ColumnHeaders.Add , , "Error", 600
    
    'add items
    Dim item As listItem
    
    For ii = 1 To count
    
        Set item = PropertiesLV.ListItems.Add
        
        item.Text = propertyIDs(ii)
        item.SubItems(1) = descriptions(ii)
        item.SubItems(2) = dataTypes(ii)
        
        If VarType(values(ii)) > vbArray Then
            item.SubItems(3) = "Array(" + CStr(UBound(values(ii))) + ")"
        Else
            item.SubItems(3) = values(ii)
        End If
        
        item.SubItems(4) = Errors(ii)

    Next ii
    
    Exit Sub
        
Failed:

End Sub
'displays the address space for the specified server.
Public Function BrowseItems(ByRef owner As Form, ByRef aServer As OPCServer, Optional ByVal isBrowsing As Boolean = True) As String()
            
On Error GoTo Failed

    'check for valid aServer object.
    If aServer Is Nothing Then
        Exit Function
    End If
    
    'save the browsing state flag.
    browseOnly = isBrowsing
    
    'initialize tabs
    If browseOnly Then
        BrowseTabs.SelectedItem = BrowseTabs.Tabs(1)
    Else
        BrowseTabs.SelectedItem = BrowseTabs.Tabs(2)
    End If
        
    'clear existing children.
    BrowserTV.Nodes.Clear
    SelectedItemsLB.Clear
            
    'create browser object.
    Set browser = aServer.CreateBrowser
    
    'fetch root branches.
    browser.MoveToRoot
    browser.ShowBranches
    
    Dim childNode As node
    Dim childElement As OPCBrowseElement
            
    Dim ii As Integer
    
    For ii = 1 To browser.count
    
        Set childElement = New OPCBrowseElement
                
        childElement.ElementName = browser.item(ii)
        childElement.ItemID = ""
        childElement.HasChildren = True
        
        Set childNode = BrowserTV.Nodes.Add(, tvwChild, CreateKey(childElement.ElementName), childElement.ElementName)
        Set childNode.Tag = childElement
        
    Next ii
    
    'fetch contents of branches.
    For ii = 1 To BrowserTV.Nodes.count
        ExpandBranch BrowserTV.Nodes.item(ii)
    Next ii
        
    'fetch root leaves.
    browser.MoveToRoot
    browser.ShowLeafs
    
    For ii = 1 To browser.count
       
        Set childElement = New OPCBrowseElement
        
        childElement.ElementName = browser.item(ii)
        childElement.ItemID = browser.GetItemID(browser.item(ii))
        childElement.HasChildren = False
        
        Set childNode = BrowserTV.Nodes.Add(, tvwChild, CreateKey(childElement.ElementName), childElement.ElementName)
        Set childNode.Tag = childElement
                
    Next ii
    
    'show the dialog.
    Set server = aServer
    Me.Show vbModal, owner
    
    'construct list of selected items for return.
    If browseOnly Or SelectedItemsLB.ListCount = 0 Then
        Exit Function
    End If
        
    Dim itemIDs() As String
    ReDim itemIDs(SelectedItemsLB.ListCount)
    
    For ii = 0 To SelectedItemsLB.ListCount - 1
        itemIDs(ii + 1) = SelectedItemsLB.List(ii)
    Next ii
    
    BrowseItems = itemIDs
    
    Exit Function
    
Failed:

    MsgBox Error

End Function

