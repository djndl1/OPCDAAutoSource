VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   Caption         =   "OPC Data Access Automation 2.02 Sample Client"
   ClientHeight    =   7785
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   14310
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox OutputLB 
      Height          =   2010
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   14055
   End
   Begin MSComctlLib.ListView UpdatesLV 
      Height          =   5415
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9551
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
   Begin MSComctlLib.TreeView ServersTV 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   9551
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      Begin VB.Menu ExitMI 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu ServerMenu 
      Caption         =   "Server"
      Begin VB.Menu AddServerMI 
         Caption         =   "Add Server..."
      End
      Begin VB.Menu AttachServerMI 
         Caption         =   "Attach Server..."
      End
      Begin VB.Menu Separator1 
         Caption         =   "-"
      End
      Begin VB.Menu BrowseItemsMI 
         Caption         =   "Browse Items..."
      End
      Begin VB.Menu AddGroupMI 
         Caption         =   "Add Group..."
      End
      Begin VB.Menu ViewStatusMI 
         Caption         =   "View Server Status..."
      End
      Begin VB.Menu RemoveServerMI 
         Caption         =   "Remove Server"
      End
   End
   Begin VB.Menu GroupMenu 
      Caption         =   "Group"
      Begin VB.Menu EditGroupMI 
         Caption         =   "Edit Group..."
      End
      Begin VB.Menu AddItemsMI 
         Caption         =   "Add Items..."
      End
      Begin VB.Menu RemoveGroupMI 
         Caption         =   "Remove Group"
      End
      Begin VB.Menu Separator2 
         Caption         =   "-"
      End
      Begin VB.Menu ReadGroupMI 
         Caption         =   "Read..."
      End
      Begin VB.Menu WriteGroupMI 
         Caption         =   "Write..."
      End
      Begin VB.Menu Separator3 
         Caption         =   "-"
      End
      Begin VB.Menu ShowUpdatesMI 
         Caption         =   "Show Updates"
         Checked         =   -1  'True
      End
      Begin VB.Menu AsyncReadMI 
         Caption         =   "Async Read..."
      End
      Begin VB.Menu AsyncWriteMI 
         Caption         =   "Async Write..."
      End
      Begin VB.Menu AsyncRefreshMI 
         Caption         =   "Async Refresh"
      End
   End
   Begin VB.Menu ItemMenu 
      Caption         =   "Item"
      Begin VB.Menu EditItemMI 
         Caption         =   "Edit Item..."
      End
      Begin VB.Menu RemoveItemMI 
         Caption         =   "Remove Item"
      End
      Begin VB.Menu Separator4 
         Caption         =   "-"
      End
      Begin VB.Menu ReadItemMI 
         Caption         =   "Read.."
      End
      Begin VB.Menu WriteItemMI 
         Caption         =   "Write..."
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============================================================================
' TITLE: MainForm.frm
'
' CONTENTS:
'
' The main window for the sample client application.
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

'the currently active group and server (i.e. callbacks are connected).
Dim WithEvents ActiveGroup As OPCGroup
Attribute ActiveGroup.VB_VarHelpID = -1
Dim WithEvents ActiveServer As OPCServer
Attribute ActiveServer.VB_VarHelpID = -1
Dim ActiveTransactionID As Long
Dim ActiveCancelID As Long

'the tree node for the currently active group.
Dim ActiveGroupNode As node
Attribute ActiveGroupNode.VB_VarHelpID = -1

'used to generated unique handles.
Dim Handle As Long

'mapping of client handles to objects (i.e. OPCItems).
Dim HandlesToItems As New Collection
'generates a unique handle.
Private Function CreateHandle() As Long
    Handle = Handle + 1
    CreateHandle = Handle
End Function
'generates a unique key.
Private Function CreateKey(ByVal prefix As String) As String
    Handle = Handle + 1
    CreateKey = prefix + " [" + CStr(Handle) + "]"
End Function
'returns the currently selected server object if any.
Private Function GetSelectedServer() As OPCServer

    Set GetSelectedServer = Nothing

    'check if anything is selected.
    If ServersTV.SelectedItem Is Nothing Then
        Exit Function
    End If
    
    'check if a server object is selected.
    If Not TypeOf ServersTV.SelectedItem.Tag Is OPCServer Then
        Exit Function
    End If
    
    'return the current selection.
    Set GetSelectedServer = ServersTV.SelectedItem.Tag

End Function
'returns the currently selected group object if any.
Private Function GetSelectedGroup() As OPCGroup

    Set GetSelectedGroup = Nothing

    'check if anything is selected.
    If ServersTV.SelectedItem Is Nothing Then
        Exit Function
    End If
    
    'check if a group object is selected.
    If Not TypeOf ServersTV.SelectedItem.Tag Is OPCGroup Then
        Exit Function
    End If
    
    'return the current selection.
    Set GetSelectedGroup = ServersTV.SelectedItem.Tag

End Function
'returns the currently selected item object if any.
Private Function GetSelectedItem() As OPCItem

    Set GetSelectedItem = Nothing

    'check if anything is selected.
    If ServersTV.SelectedItem Is Nothing Then
        Exit Function
    End If
    
    'check if a item object is selected.
    If Not TypeOf ServersTV.SelectedItem.Tag Is OPCItem Then
        Exit Function
    End If
    
    'return the current selection.
    Set GetSelectedItem = ServersTV.SelectedItem.Tag

End Function
'enables menus based on the current selection.
Private Sub EnableMenus()
            
        BrowseItemsMI.Enabled = False
        AddGroupMI.Enabled = False
        ViewStatusMI.Enabled = False
        RemoveServerMI.Enabled = False
        AddItemsMI.Enabled = False
        EditGroupMI.Enabled = False
        RemoveGroupMI.Enabled = False
        ReadGroupMI.Enabled = False
        WriteGroupMI.Enabled = False
        ShowUpdatesMI.Enabled = False
        ShowUpdatesMI.Checked = False
        AsyncReadMI.Enabled = False
        AsyncWriteMI.Enabled = False
        AsyncRefreshMI.Enabled = False
        EditItemMI.Enabled = False
        RemoveItemMI.Enabled = False
        ReadItemMI.Enabled = False
        WriteItemMI.Enabled = False
        
        'handle item selected.
        If Not GetSelectedItem() Is Nothing Then
        
            EditItemMI.Enabled = True
            RemoveItemMI.Enabled = True
            ReadItemMI.Enabled = True
            WriteItemMI.Enabled = True
            
            Exit Sub
        End If
                        
        'handle group selected.
        Dim group As OPCGroup
        Set group = GetSelectedGroup()
        
        If Not group Is Nothing Then
            
            AddItemsMI.Enabled = True
            EditGroupMI.Enabled = True
            RemoveGroupMI.Enabled = True
            ReadGroupMI.Enabled = True
            WriteGroupMI.Enabled = True
            
            If group.IsSubscribed Then
                ShowUpdatesMI.Enabled = True
            
                If group Is ActiveGroup Then
                    ShowUpdatesMI.Checked = True
                    AsyncReadMI.Enabled = True
                    AsyncWriteMI.Enabled = True
                    AsyncRefreshMI.Enabled = True
                End If
            End If
         
            Exit Sub
        End If

        'handle server selected.
        If Not GetSelectedServer() Is Nothing Then
        
            BrowseItemsMI.Enabled = True
            AddGroupMI.Enabled = True
            ViewStatusMI.Enabled = True
            RemoveServerMI.Enabled = True
            
            Exit Sub
        End If
        
End Sub
'handles a cancel complete transaction.
Private Sub ActiveGroup_AsyncCancelComplete(ByVal CancelID As Long)
    
    'update output.
    OutputLB.AddItem "Cancel #" + CStr(CancelID) + " Complete.", 1
    
    'check transaction id.
    If Not CancelID = ActiveTransactionID Then
        Exit Sub
    End If
    
    ActiveTransactionID = 0
    ActiveCancelID = 0
    
End Sub

'handles read callbacks for active group.
Private Sub ActiveGroup_AsyncReadComplete(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, ItemValues() As Variant, Qualities() As Long, TimeStamps() As Date, Errors() As Long)
    
    'update output.
    OutputLB.AddItem "Read #" + CStr(TransactionID) + " Complete.", 0
    
    'check transaction id.
    If Not TransactionID = ActiveTransactionID Then
        Exit Sub
    End If
    
    ActiveTransactionID = 0
    ActiveCancelID = 0
    
    'do nothing if no active group.
    If ActiveGroup Is Nothing Then
        Exit Sub
    End If
       
    'lookup items
    Dim items() As OPCItem
    ReDim items(NumItems)
    
    Dim ii As Integer
    
    For ii = 1 To NumItems
        Set items(ii) = HandlesToItems.item(CStr(ClientHandles(ii)))
    Next ii
       
    'display the results
    ItemValuesDlg.ShowItems Me, NumItems, items, ItemValues, Qualities, TimeStamps, Errors
    
End Sub
'handles write callbacks for active group.
Private Sub ActiveGroup_AsyncWriteComplete(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, Errors() As Long)
    
    'update output.
    OutputLB.AddItem "Write #" + CStr(TransactionID) + " Complete.", 1
    
    'check transaction id.
    If Not TransactionID = ActiveTransactionID Then
        Exit Sub
    End If
    
    ActiveTransactionID = 0
    ActiveCancelID = 0
    
    'do nothing if no active group.
    If ActiveGroup Is Nothing Then
        Exit Sub
    End If
    
    'lookup items
    Dim items() As OPCItem
    Dim values() As Variant
    ReDim items(NumItems)
    ReDim values(NumItems)
    
    Dim ii As Integer
    
    For ii = 1 To NumItems
        Set items(ii) = HandlesToItems.item(CStr(ClientHandles(ii)))
        values(ii) = items(ii).value
    Next ii
    
    'display any errors.
    ItemValuesDlg.ShowErrors Me, NumItems, items, values, Errors

End Sub
'handles callbacks for active group.
Private Sub ActiveGroup_DataChange(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, ItemValues() As Variant, Qualities() As Long, TimeStamps() As Date)

    'check for refresh callbacks.
    If Not TransactionID = 0 Then
    
        'update output
        OutputLB.AddItem "Refresh #" + CStr(TransactionID) + " Complete.", 1
        
        'check transaction id.
        If Not TransactionID = ActiveTransactionID Then
            Exit Sub
        End If
        
        ActiveTransactionID = 0
        ActiveCancelID = 0
    
    End If
    
    'do nothing if no active group.
    If ActiveGroup Is Nothing Then
        Exit Sub
    End If
    
    'do nothing if no active group.
    If ActiveGroup Is Nothing Then
        Exit Sub
    End If

    'clear current list contents.
    UpdatesLV.ListItems.Clear
    
    'update list.
    Dim ii As Integer
    Dim item As OPCItem
    Dim listItem As listItem
    
    For ii = 1 To NumItems
    
        Set listItem = UpdatesLV.ListItems.Add
        
        listItem.Text = ActiveGroup.Name
        
        Set item = HandlesToItems.item(CStr(ClientHandles(ii)))
        
        If Not item Is Nothing Then
            listItem.SubItems(1) = item.ItemID
        End If
        
        If VarType(ItemValues(ii)) > vbArray Then
            listItem.SubItems(2) = "Array(" + CStr(UBound(ItemValues(ii)) - 1) + ")"
        Else
            listItem.SubItems(2) = ItemValues(ii)
        End If
        
        listItem.SubItems(3) = Qualities(ii)
        listItem.SubItems(4) = TimeStamps(ii)

    Next ii
        
End Sub
'called when the active server shutsdown.
Private Sub ActiveServer_ServerShutDown(ByVal Reason As String)
    MsgBox "Server shutting down. Reason = " + Reason
End Sub

Private Sub AttachServerMI_Click()
    
On Error GoTo Failed

    'select the server.
    Dim progID As String
    progID = SelectServerDlg.SelectServer(Me, "")
    
    If progID = "" Then
        Exit Sub
    End If
    
    'create a new server object.
    Dim customServer
    Set customServer = CreateObject(progID)
    
    Dim activator As OPCActivator
    Set activator = New OPCActivator
    
    'create wrapper
    Dim server As OPCServer
    Set server = activator.Attach(customServer, progID)
        
    server.ClientName = CreateKey(server.ServerName)
    
    'update tree view
    Dim child As node
    Set child = ServersTV.Nodes.Add(, tvwChild, server.ClientName, server.ServerName)
    Set child.Tag = server
        
    ServersTV.SelectedItem = child
    
    Exit Sub
    
Failed:
    
    MsgBox Error
    
End Sub

'closes the form.
Private Sub ExitMI_Click()
    Unload Me
End Sub
'intializes the form on load.
Private Sub Form_Load()
    
    EnableMenus
    
    UpdatesLV.ColumnHeaders.Add , , "Group", 1000
    UpdatesLV.ColumnHeaders.Add , , "Item", 2500
    UpdatesLV.ColumnHeaders.Add , , "Value", 2000
    UpdatesLV.ColumnHeaders.Add , , "Quality", 1000
    UpdatesLV.ColumnHeaders.Add , , "Timestamp", 2000
    
End Sub
'cleans up on exit.
Private Sub Form_Unload(Cancel As Integer)
    
    'remove all groups and disconnect from all servers.
    Dim ii As Integer
    Dim server As OPCServer
    
    For ii = 1 To ServersTV.Nodes.count
        
        'check for server node.
        If TypeOf ServersTV.Nodes(ii).Tag Is OPCServer Then
            Set server = ServersTV.Nodes(ii).Tag
            server.OPCGroups.RemoveAll
            server.Disconnect
            Set server = Nothing
        End If
        
        'remove reference to object.
        Set ServersTV.Nodes(ii).Tag = Nothing
    
    Next ii
    
    'clear the tree view.
    ServersTV.Nodes.Clear
        
End Sub
'enables menus based on the current selection.
Private Sub ServersTV_Click()
    EnableMenus
End Sub

'displays the popup menu in the servers tree view.
Private Sub ServersTV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbRightButton Then
        
        'enable menus based on current selection.
        EnableMenus
        
        'handle item selected.
        If Not GetSelectedItem() Is Nothing Then
            PopupMenu ItemMenu
            Exit Sub
        End If
                        
        'handle group selected.
        If Not GetSelectedGroup() Is Nothing Then
            PopupMenu GroupMenu
            Exit Sub
        End If

        'handle server selected.
        If Not GetSelectedServer() Is Nothing Then
            PopupMenu ServerMenu
            Exit Sub
        End If

        'handle nothing selected.
        PopupMenu ServerMenu
        
    End If

End Sub
'selects and connects to a new OPC server.
Private Sub AddServerMI_Click()
    
On Error GoTo Failed

    'select the server.
    Dim progID As String
    progID = SelectServerDlg.SelectServer(Me, "")
    
    If progID = "" Then
        Exit Sub
    End If
    
    'create a new server object.
    Dim server As OPCServer
    Set server = New OPCServer
    
    'connect to the server.
    server.Connect (progID)
    server.ClientName = CreateKey(server.ServerName)
    
    'update tree view
    Dim child As node
    Set child = ServersTV.Nodes.Add(, tvwChild, server.ClientName, server.ServerName)
    Set child.Tag = server
        
    ServersTV.SelectedItem = child
    
    Exit Sub
    
Failed:
    
    MsgBox Error
    
End Sub
'removes the currently selected OPC server.
Private Sub RemoveServerMI_Click()

On Error GoTo Failed
    
    'find the currently selected server object.
    Dim server As OPCServer
    Set server = GetSelectedServer()
    
    If server Is Nothing Then
        Exit Sub
    End If
    
    Dim key As String
    key = server.ClientName
    
    'disconnect from server
    server.Disconnect

Failed:
            
    'remove server from display.
    ServersTV.Nodes.Remove server.ClientName

End Sub
'shows the address space for the current server.
Private Sub BrowseItemsMI_Click()
    
    'find the currently selected server object.
    Dim server As OPCServer
    Set server = GetSelectedServer()
    
    If server Is Nothing Then
        Exit Sub
    End If
    
    'show the browse items dialog
    BrowseItemsDlg.BrowseItems Me, server
        
End Sub
'changes the active group.
Private Sub ShowUpdatesMI_Click()

On Error GoTo Failed

    'get currently selected group.
    Dim group As OPCGroup
    Set group = GetSelectedGroup()
        
    If group Is Nothing Then
        Exit Sub
    End If
    
    'turn off active group.
    If ShowUpdatesMI.Checked Then
    
        ShowUpdatesMI.Checked = False
        
        If group Is ActiveGroup Then
            Set ActiveGroup = Nothing
            Set ActiveServer = Nothing
            
            If Not ActiveGroupNode Is Nothing Then
                ActiveGroupNode.ForeColor = vbWindowText
                ActiveGroupNode.parent.ForeColor = vbWindowText
                Set ActiveGroupNode = Nothing
            End If
            
        End If
    
    'turn on new group
    Else
    
        If group.IsSubscribed Then
            ShowUpdatesMI.Checked = True
            Set ActiveGroup = group
            Set ActiveServer = group.parent
                        
            If Not ActiveGroupNode Is Nothing Then
                ActiveGroupNode.ForeColor = vbWindowText
                ActiveGroupNode.parent.ForeColor = vbWindowText
            End If
            
            Set ActiveGroupNode = ServersTV.SelectedItem
            ActiveGroupNode.ForeColor = &HFF0000
            ActiveGroupNode.parent.ForeColor = &HFF0000
        End If
    
    End If
    
    Exit Sub
    
Failed:

    MsgBox Error
    
End Sub
'shows the properties of the current server.
Private Sub ViewStatusMI_Click()
    
    'find the currently selected server object.
    Dim server As OPCServer
    Set server = GetSelectedServer()
    
    If server Is Nothing Then
        Exit Sub
    End If
    
    'show the browse items dialog
    ViewStatusDlg.ViewStatus Me, server

End Sub
'adds a group to the server.
Private Sub AddGroupMI_Click()
    
On Error GoTo Failed

    'find the currently selected server object.
    Dim server As OPCServer
    Set server = GetSelectedServer()
    
    If server Is Nothing Then
        Exit Sub
    End If
    
    'create default group.
    Dim group As OPCGroup
    Set group = server.OPCGroups.Add("")
    
    'turn subscriptions on if currently active.
    If group.IsActive Then
        group.IsSubscribed = True
    End If
    
    'edit group parameters
    Dim result As Boolean
    result = EditGroupDlg.EditGroup(Me, group)
    
    If Not result Then
        server.OPCGroups.Remove (group.Name)
        Exit Sub
    End If
    
    'save group parameters as default.
    server.OPCGroups.DefaultGroupIsActive = group.IsActive
    server.OPCGroups.DefaultGroupLocaleID = group.LocaleID
    server.OPCGroups.DefaultGroupTimeBias = group.TimeBias
    server.OPCGroups.DefaultGroupDeadband = group.DeadBand
    server.OPCGroups.DefaultGroupUpdateRate = group.UpdateRate
    
    'update tree view
    Dim child As node
    Set child = ServersTV.Nodes.Add(ServersTV.SelectedItem, tvwChild, CreateKey(group.Name), group.Name)
    Set child.Tag = group
    child.EnsureVisible
    
    ServersTV.SelectedItem = child
    
    'set active group is not already set.
    If ActiveGroup Is Nothing Then
        Set ActiveGroup = group
        Set ActiveServer = server
        Set ActiveGroupNode = child
        ActiveGroupNode.ForeColor = &HFF0000
        ActiveGroupNode.parent.ForeColor = &HFF0000
    End If
       
    Exit Sub
         
Failed:
    
    MsgBox Error
            
End Sub
'edits an existing group.
Private Sub EditGroupMI_Click()

On Error GoTo Failed

    'find the currently selected group object.
    Dim group As OPCGroup
    Set group = GetSelectedGroup()
    
    If group Is Nothing Then
        Exit Sub
    End If
    
    'edit group parameters
    Dim result As Boolean
    result = EditGroupDlg.EditGroup(Me, group)
    
    If Not result Then
        Exit Sub
    End If
            
    'update tree view
    ServersTV.SelectedItem.Text = group.Name
          
    Exit Sub
         
Failed:
    
    MsgBox Error

End Sub
'removes an existing group.
Private Sub RemoveGroupMI_Click()

On Error GoTo Failed

    'find the currently selected group object.
    Dim group As OPCGroup
    Set group = GetSelectedGroup()
    
    If group Is Nothing Then
        Exit Sub
    End If
    
    'remove group from server.
    Dim server As OPCServer
    Set server = group.parent
    
    server.OPCGroups.Remove group.Name
            
    If ActiveGroup Is group Then
        Set ActiveGroup = Nothing
        Set ActiveServer = Nothing
    End If
    
    'update tree view
    Dim child As node
    Set child = ServersTV.SelectedItem
    
    If ActiveGroupNode Is child Then
        Set ActiveGroupNode = Nothing
    End If
    
    ServersTV.SelectedItem = child.Previous
    ServersTV.Nodes.Remove child.key
          
    Exit Sub
         
Failed:
    
    MsgBox Error

End Sub
'prompts the user to select items to add.
Private Sub AddItemsMI_Click()
    
On Error GoTo Failed

    'find the currently selected group object.
    Dim group As OPCGroup
    Set group = GetSelectedGroup()
    
    If group Is Nothing Then
        Exit Sub
    End If
            
    'select items.
    Dim itemIDs() As String
    itemIDs = BrowseItemsDlg.BrowseItems(Me, group.parent, False)
    
    'check if anything was actually selected.
    If UBound(itemIDs) = 0 Then
        Exit Sub
    End If
    
    'assign unique client handles.
    Dim ii As Integer
    Dim ClientHandles() As Long
    ReDim ClientHandles(UBound(itemIDs))
    
    For ii = 1 To UBound(itemIDs)
        ClientHandles(ii) = CreateHandle()
    Next ii
                 
    'add items to group.
    Dim serverHandles() As Long
    Dim Errors() As Long
    
    group.OPCItems.AddItems UBound(itemIDs), itemIDs, ClientHandles, serverHandles, Errors
    
    'update tree view.
    Dim child As node
    Dim newItem As OPCItem
        
    For ii = 1 To UBound(itemIDs)
    
        'check for error.
        If Errors(ii) = 0 Then
            Set newItem = group.OPCItems.item(itemIDs(ii))
            HandlesToItems.Add newItem, CStr(ClientHandles(ii))
            
            Set child = ServersTV.Nodes.Add(ServersTV.SelectedItem, tvwChild, CreateKey(itemIDs(ii)), itemIDs(ii))
            Set child.Tag = newItem
        End If
        
    Next ii
        
    child.EnsureVisible
    ServersTV.SelectedItem = child
                 
    Exit Sub
         
Failed:
    
    MsgBox Error

End Sub
'edits an existing item.
Private Sub EditItemMI_Click()

On Error GoTo Failed

    'find the currently selected group object.
    Dim item As OPCItem
    Set item = GetSelectedItem()
    
    If item Is Nothing Then
        Exit Sub
    End If
    
    'edit item parameters
    EditItemDlg.EditItem Me, item
          
    Exit Sub
         
Failed:
    
    MsgBox Error

End Sub
'removes an existing item.
Private Sub RemoveItemMI_Click()

On Error GoTo Failed

    'find the currently selected item object.
    Dim item As OPCItem
    Set item = GetSelectedItem()
    
    If item Is Nothing Then
        Exit Sub
    End If
        
    'remove item from group.
    Dim group As OPCGroup
    Set group = item.parent
    
    Dim serverHandles(1) As Long
    serverHandles(1) = item.ServerHandle
         
    Dim Errors() As Long
    
    group.OPCItems.Remove 1, serverHandles, Errors
    
    'update tree view
    Dim child As node
    Set child = ServersTV.SelectedItem
    ServersTV.SelectedItem = child.Previous
    ServersTV.Nodes.Remove child.key
          
    Exit Sub
         
Failed:
    
    MsgBox Error

End Sub
'reads the current values for the selected group.
Private Sub ReadGroupMI_Click()

On Error GoTo Failed

    'find the currently selected group object.
    Dim group As OPCGroup
    Set group = GetSelectedGroup()
    
    If group Is Nothing Then
        Exit Sub
    End If
    
    'check if items exist.
    Dim count As Long
    count = group.OPCItems.count
    
    If count = 0 Then
        Exit Sub
    End If
    
    'set the item server handles.
    Dim items() As OPCItem
    Dim serverHandles() As Long
    
    ReDim items(count)
    ReDim serverHandles(count)
        
    Dim ii As Long
    
    For ii = 1 To count
        Set items(ii) = group.OPCItems(ii)
        serverHandles(ii) = group.OPCItems(ii).ServerHandle
    Next ii
    
    'read the values.
    Dim values() As Variant
    Dim Errors() As Long
    Dim quality As Variant
    Dim timestamp As Variant
       
    group.SyncRead OPCDataSource.OPCDevice, count, serverHandles, values, Errors, quality, timestamp
        
    'convert variants to arrays.
    Dim Qualities() As Long
    ReDim Qualities(count)
    
    If VarType(quality) = vbArray + vbInteger Then
        
        Dim buffer() As Integer
        buffer = quality
        
        For ii = 1 To count
            Qualities(ii) = buffer(ii)
        Next ii
            
    End If
    
    Dim TimeStamps() As Date
    If VarType(timestamp) = vbArray + vbDate Then
        TimeStamps = timestamp
    End If
    
    'display the results
    ItemValuesDlg.ShowItems Me, count, items, values, Qualities, TimeStamps, Errors
    
    Exit Sub
         
Failed:
    
    MsgBox Error
    
End Sub
'reads the current value for the selected item.
Private Sub ReadItemMI_Click()

On Error GoTo Failed

    'find the currently selected group object.
    Dim items(1) As OPCItem
    Set items(1) = GetSelectedItem()
    
    If items(1) Is Nothing Then
        Exit Sub
    End If
    
    'read the item value.
    Dim values(1) As Variant
    Dim quality As Variant
    Dim timestamp As Variant
    Dim Errors(1) As Long
       
    items(1).Read OPCDataSource.OPCDevice, values(1), quality, timestamp
    
    'convert variants to arrays.
    Dim Qualities(1) As Long
    If VarType(quality) = vbInteger Then
        Qualities(1) = quality
    End If
    
    Dim TimeStamps(1) As Date
    If VarType(timestamp) = vbDate Then
        TimeStamps(1) = timestamp
    End If
    
    'display the results
    ItemValuesDlg.ShowItems Me, 1, items, values, Qualities, TimeStamps, Errors
    
    Exit Sub
         
Failed:
    
    MsgBox Error
    
End Sub
'writes the current values for the selected group.
Private Sub WriteGroupMI_Click()

On Error GoTo Failed

    'find the currently selected group object.
    Dim group As OPCGroup
    Set group = GetSelectedGroup()
    
    If group Is Nothing Then
        Exit Sub
    End If
    
    'check if items exist.
    Dim count As Long
    count = group.OPCItems.count
    
    If count = 0 Then
        Exit Sub
    End If
    
    'set the item server handles.
    Dim items() As OPCItem
    Dim serverHandles() As Long
    
    ReDim items(count)
    ReDim serverHandles(count)
        
    Dim ii As Long
    
    For ii = 1 To count
        Set items(ii) = group.OPCItems(ii)
        serverHandles(ii) = group.OPCItems(ii).ServerHandle
    Next ii
    
    'read the current values.
    Dim values() As Variant
    Dim Errors() As Long
       
    group.SyncRead OPCDataSource.OPCDevice, count, serverHandles, values, Errors
    
    'write back the current values.
    group.SyncWrite count, serverHandles, values, Errors
        
    'display any errors.
    ItemValuesDlg.ShowErrors Me, count, items, values, Errors
    
    Exit Sub
         
Failed:
    
    MsgBox Error
    
End Sub
'reads the current value for the selected item.
Private Sub WriteItemMI_Click()

On Error GoTo Failed

    'find the currently selected group object.
    Dim item As OPCItem
    Set item = GetSelectedItem()
    
    If item Is Nothing Then
        Exit Sub
    End If
    
    'read the current item value.
    Dim value As Variant
    item.Read OPCDataSource.OPCDevice, value
    
    'prompt user to edit value.
    value = EditValueDlg.EditValue(Me, value)
    
    If value = Empty Then
        Exit Sub
    End If
        
    'write new value.
    item.Write value
    
    Exit Sub
         
Failed:
    
    MsgBox Error
    
End Sub
'reads the current values for the selected group.
Private Sub AsyncReadMI_Click()

On Error GoTo Failed

    'find the currently selected group object.
    Dim group As OPCGroup
    Set group = GetSelectedGroup()
    
    If group Is Nothing Or Not group Is ActiveGroup Then
        Exit Sub
    End If
    
    'check if items exist.
    Dim count As Long
    count = group.OPCItems.count
    
    If count = 0 Then
        Exit Sub
    End If
    
    'set the item server handles.
    Dim items() As OPCItem
    Dim serverHandles() As Long
    
    ReDim items(count)
    ReDim serverHandles(count)
        
    Dim ii As Long
    
    For ii = 1 To count
        Set items(ii) = group.OPCItems(ii)
        serverHandles(ii) = group.OPCItems(ii).ServerHandle
    Next ii
    
On Error GoTo SkipCancel

    'cancel existing transaction.
    If Not ActiveTransactionID = 0 Then
        group.AsyncCancel ActiveCancelID
    End If

SkipCancel:
On Error GoTo Failed
    
    'read the values.
    Dim Errors() As Long
       
    'begin read transaction.
    ActiveTransactionID = CreateHandle()
    group.AsyncRead count, serverHandles, Errors, ActiveTransactionID, ActiveCancelID
    
    'update output
    OutputLB.AddItem "Read #" + CStr(ActiveTransactionID) + " Started.", 0
    
    Exit Sub
         
Failed:
    
    MsgBox Error
    
End Sub
'writes the current values for the selected group.
Private Sub AsyncWriteMI_Click()

On Error GoTo Failed

    'find the currently selected group object.
    Dim group As OPCGroup
    Set group = GetSelectedGroup()
    
    If Not Not group Is Nothing Or Not group Is ActiveGroup Then
        Exit Sub
    End If
    
    'check if items exist.
    Dim count As Long
    count = group.OPCItems.count
    
    If count = 0 Then
        Exit Sub
    End If
    
    'set the item server handles.
    Dim items() As OPCItem
    Dim serverHandles() As Long
    
    ReDim items(count)
    ReDim serverHandles(count)
        
    Dim ii As Long
    
    For ii = 1 To count
        Set items(ii) = group.OPCItems(ii)
        serverHandles(ii) = group.OPCItems(ii).ServerHandle
    Next ii
    
    'read the current values.
    Dim values() As Variant
    Dim Errors() As Long
       
    group.SyncRead OPCDataSource.OPCDevice, count, serverHandles, values, Errors
    
On Error GoTo SkipCancel

    'cancel any active transaction.
    If Not ActiveTransactionID = 0 Then
        group.AsyncCancel ActiveCancelID
    End If
           
SkipCancel:
On Error GoTo Failed

    'begin write transaction.
    ActiveTransactionID = CreateHandle()
    group.AsyncWrite count, serverHandles, values, Errors, ActiveTransactionID, ActiveCancelID
    
    'update output
    OutputLB.AddItem "Write #" + CStr(ActiveTransactionID) + " Started.", 0
    
    Exit Sub
         
Failed:
    
    MsgBox Error
    
End Sub
'refreshes current values for the selected group.
Private Sub AsyncRefreshMI_Click()

On Error GoTo Failed

    'find the currently selected group object.
    Dim group As OPCGroup
    Set group = GetSelectedGroup()
    
    If Not Not group Is Nothing Or Not group Is ActiveGroup Then
        Exit Sub
    End If
    
    'check if items exist.
    Dim count As Long
    count = group.OPCItems.count
    
    If count = 0 Then
        Exit Sub
    End If
    
On Error GoTo SkipCancel

    'cancel any active transaction.
    If Not ActiveTransactionID = 0 Then
        group.AsyncCancel ActiveCancelID
    End If
           
SkipCancel:
On Error GoTo Failed

    'begin refresh transaction.
    ActiveTransactionID = CreateHandle()
    group.AsyncRefresh OPCDataSource.OPCDevice, ActiveTransactionID, ActiveCancelID
    
    'update output
    OutputLB.AddItem "Refresh #" + CStr(ActiveTransactionID) + " Started.", 0
    
    Exit Sub
         
Failed:
    
    MsgBox Error
    
End Sub

