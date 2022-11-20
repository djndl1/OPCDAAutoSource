'============================================================================
' TITLE: GroupsCtrl.vb
'
' CONTENTS:
' 
' A control used to manage a set of groups for a server.
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

Imports OPCAutomation

Public Class GroupsCtrl
    Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call       
        Static Resources As Resources = New Resources
        GroupsTV.ImageList = Resources.ImageList

    End Sub

    'UserControl overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents GroupsTV As System.Windows.Forms.TreeView
    Friend WithEvents PopupMenu As System.Windows.Forms.ContextMenu
    Friend WithEvents AddGroupMI As System.Windows.Forms.MenuItem
    Friend WithEvents EditGroupMI As System.Windows.Forms.MenuItem
    Friend WithEvents RemoveGroupMI As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents AddItemsMI As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents SyncReadMI As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents AsyncReadMI As System.Windows.Forms.MenuItem
    Friend WithEvents SyncWriteMI As System.Windows.Forms.MenuItem
    Friend WithEvents AsyncWriteMI As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents BrowseItemsMI As System.Windows.Forms.MenuItem
    Friend WithEvents AsyncRefreshMI As System.Windows.Forms.MenuItem
    Friend WithEvents RemoveItemMI As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents EditItemMI As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupsTV = New System.Windows.Forms.TreeView
        Me.PopupMenu = New System.Windows.Forms.ContextMenu
        Me.AddGroupMI = New System.Windows.Forms.MenuItem
        Me.BrowseItemsMI = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.EditGroupMI = New System.Windows.Forms.MenuItem
        Me.RemoveGroupMI = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.AddItemsMI = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.SyncReadMI = New System.Windows.Forms.MenuItem
        Me.SyncWriteMI = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.AsyncReadMI = New System.Windows.Forms.MenuItem
        Me.AsyncWriteMI = New System.Windows.Forms.MenuItem
        Me.AsyncRefreshMI = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.RemoveItemMI = New System.Windows.Forms.MenuItem
        Me.EditItemMI = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'GroupsTV
        '
        Me.GroupsTV.ContextMenu = Me.PopupMenu
        Me.GroupsTV.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupsTV.ImageIndex = -1
        Me.GroupsTV.Location = New System.Drawing.Point(0, 0)
        Me.GroupsTV.Name = "GroupsTV"
        Me.GroupsTV.SelectedImageIndex = -1
        Me.GroupsTV.Size = New System.Drawing.Size(360, 392)
        Me.GroupsTV.TabIndex = 0
        '
        'PopupMenu
        '
        Me.PopupMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.AddGroupMI, Me.BrowseItemsMI, Me.MenuItem3, Me.EditGroupMI, Me.RemoveGroupMI, Me.MenuItem1, Me.AddItemsMI, Me.MenuItem4, Me.SyncReadMI, Me.SyncWriteMI, Me.MenuItem6, Me.AsyncReadMI, Me.AsyncWriteMI, Me.AsyncRefreshMI, Me.MenuItem5, Me.EditItemMI, Me.RemoveItemMI})
        '
        'AddGroupMI
        '
        Me.AddGroupMI.Index = 0
        Me.AddGroupMI.Text = "Add Group..."
        '
        'BrowseItemsMI
        '
        Me.BrowseItemsMI.Index = 1
        Me.BrowseItemsMI.Text = "Browse Items..."
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 2
        Me.MenuItem3.Text = "-"
        '
        'EditGroupMI
        '
        Me.EditGroupMI.Index = 3
        Me.EditGroupMI.Text = "Edit Group..."
        '
        'RemoveGroupMI
        '
        Me.RemoveGroupMI.Index = 4
        Me.RemoveGroupMI.Text = "Remove Group"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 5
        Me.MenuItem1.Text = "-"
        '
        'AddItemsMI
        '
        Me.AddItemsMI.Index = 6
        Me.AddItemsMI.Text = "Add Items..."
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 7
        Me.MenuItem4.Text = "-"
        '
        'SyncReadMI
        '
        Me.SyncReadMI.Index = 8
        Me.SyncReadMI.Text = "Sync Read..."
        '
        'SyncWriteMI
        '
        Me.SyncWriteMI.Index = 9
        Me.SyncWriteMI.Text = "Sync Write..."
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 10
        Me.MenuItem6.Text = "-"
        '
        'AsyncReadMI
        '
        Me.AsyncReadMI.Index = 11
        Me.AsyncReadMI.Text = "Async Read..."
        '
        'AsyncWriteMI
        '
        Me.AsyncWriteMI.Index = 12
        Me.AsyncWriteMI.Text = "Async Write..."
        '
        'AsyncRefreshMI
        '
        Me.AsyncRefreshMI.Index = 13
        Me.AsyncRefreshMI.Text = "Async Refresh"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 14
        Me.MenuItem5.Text = "-"
        '
        'RemoveItemMI
        '
        Me.RemoveItemMI.Index = 16
        Me.RemoveItemMI.Text = "Remove Item"
        '
        'EditItemMI
        '
        Me.EditItemMI.Index = 15
        Me.EditItemMI.Text = "Edit Item..."
        '
        'GroupsCtrl
        '
        Me.Controls.Add(Me.GroupsTV)
        Me.Name = "GroupsCtrl"
        Me.Size = New System.Drawing.Size(360, 392)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private m_aServer As Server

    Public Delegate Sub GroupSubscribed_EventHandler(ByRef pGroup As OPCGroup)

    Public Event GroupSubscribed As GroupSubscribed_EventHandler

    Public Sub Initailize(ByRef server As Server)

        m_aServer = server

        If m_aServer Is Nothing Then
            GroupsTV.Nodes.Clear()
            Exit Sub
        End If

        Dim node As TreeNode = New TreeNode(server.Server.ServerName)

        node.ImageIndex = Resources.SERVER
        node.SelectedImageIndex = Resources.SERVER
        node.Tag = server.Server

        GroupsTV.Nodes.Add(node)
        node.EnsureVisible()

    End Sub

    Public Sub AddGroup(ByRef parent As TreeNode, ByRef pGroup As OPCGroup)

        m_aServer.AddGroup(pGroup)

        Dim node As TreeNode = New TreeNode(pGroup.Name)

        node.ImageIndex = Resources.ALTERNATE_FOLDER
        node.SelectedImageIndex = Resources.ALTERNATE_FOLDER
        node.Tag = pGroup

        parent.Nodes.Add(node)
        node.EnsureVisible()

    End Sub
    Public Sub AddItems(ByRef parent As TreeNode, ByRef pGroup As OPCGroup)

        If Not TypeOf (parent.Tag) Is OPCGroup Then
            Exit Sub
        End If

        Dim group As Group = m_aServer.GetGroup(pGroup)

        parent.Nodes.Clear()

        Dim pItems As OPCItems = pGroup.OPCItems

        For Each pItem As OPCItem In pItems

            Dim item As Item = group.GetItem(pItem)

            If Not item Is Nothing Then
                Dim node As TreeNode = New TreeNode(item.ItemID)

                node.ImageIndex = Resources.ITEM
                node.SelectedImageIndex = Resources.ITEM
                node.Tag = pItem

                parent.Nodes.Add(node)
                node.EnsureVisible()
            End If

        Next

    End Sub

    Private Sub GroupsTV_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles GroupsTV.MouseDown

        If e.Button <> MouseButtons.Right Then
            Exit Sub
        End If

        AddGroupMI.Enabled = False
        BrowseItemsMI.Enabled = False
        EditGroupMI.Enabled = False
        RemoveGroupMI.Enabled = False
        AddItemsMI.Enabled = False
        SyncReadMI.Enabled = False
        AsyncReadMI.Enabled = False
        SyncWriteMI.Enabled = False
        AsyncWriteMI.Enabled = False
        AsyncRefreshMI.Enabled = False
        EditItemMI.Enabled = False
        RemoveItemMI.Enabled = False

        Dim clickedNode As TreeNode = GroupsTV.GetNodeAt(e.X, e.Y)

        If clickedNode Is Nothing Then
            Exit Sub
        End If

        GroupsTV.SelectedNode = clickedNode

        If TypeOf (clickedNode.Tag) Is OPCServer Then
            AddGroupMI.Enabled = True
            BrowseItemsMI.Enabled = True
        End If

        If TypeOf (clickedNode.Tag) Is OPCGroup Then
            EditGroupMI.Enabled = True
            RemoveGroupMI.Enabled = True
            AddItemsMI.Enabled = True

            If clickedNode.Nodes.Count > 0 Then
                SyncReadMI.Enabled = True
                SyncWriteMI.Enabled = True

                Dim pGroup As OPCGroup = DirectCast(clickedNode.Tag, OPCGroup)

                If pGroup.IsSubscribed Then
                    AsyncReadMI.Enabled = True
                    AsyncWriteMI.Enabled = True
                    AsyncRefreshMI.Enabled = True
                End If
            End If
        End If

        If TypeOf (clickedNode.Tag) Is OPCItem Then
            EditItemMI.Enabled = True
            RemoveItemMI.Enabled = True
        End If

    End Sub

    Private Sub AddGroupMI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AddGroupMI.Click

        Dim node As TreeNode = GroupsTV.SelectedNode

        If node Is Nothing Or Not TypeOf (node.Tag) Is OPCServer Then
            Exit Sub
        End If

        Dim pGroup As OPCGroup = New AddGroupDlg().ShowDialog(DirectCast(node.Tag, OPCServer))

        If pGroup Is Nothing Then
            Exit Sub
        End If

        AddGroup(node, pGroup)
        RaiseEvent GroupSubscribed(pGroup)

    End Sub

    Private Sub EditGroupMI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles EditGroupMI.Click

        Dim node As TreeNode = GroupsTV.SelectedNode

        If node Is Nothing Or Not TypeOf (node.Tag) Is OPCGroup Then
            Exit Sub
        End If

        Dim pGroup As OPCGroup = DirectCast(node.Tag, OPCGroup)

        If Not New EditGroupDlg().ShowDialog(pGroup) Then
            Exit Sub
        End If

        node.Text = pGroup.Name
        RaiseEvent GroupSubscribed(pGroup)

    End Sub

    Private Sub AddItemsMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddItemsMI.Click

        Dim node As TreeNode = GroupsTV.SelectedNode

        If node Is Nothing Or Not TypeOf (node.Tag) Is OPCGroup Then
            Exit Sub
        End If

        Dim pGroup As OPCGroup = DirectCast(node.Tag, OPCGroup)

        If Not New AddItemsDlg().ShowDialog(m_aServer, pGroup) Then
            Exit Sub
        End If

        AddItems(node, pGroup)

    End Sub

    Private Sub SyncReadMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SyncReadMI.Click

        Dim node As TreeNode = GroupsTV.SelectedNode

        If node Is Nothing Or Not TypeOf (node.Tag) Is OPCGroup Then
            Exit Sub
        End If

        Dim pGroup As OPCGroup = DirectCast(node.Tag, OPCGroup)

        If Not New ReadItemsDlg().ShowDialog(m_aServer, pGroup, False) Then
            Exit Sub
        End If

    End Sub

    Private Sub AsyncReadMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AsyncReadMI.Click

        Dim node As TreeNode = GroupsTV.SelectedNode

        If node Is Nothing Or Not TypeOf (node.Tag) Is OPCGroup Then
            Exit Sub
        End If

        Dim pGroup As OPCGroup = DirectCast(node.Tag, OPCGroup)

        If Not New ReadItemsDlg().ShowDialog(m_aServer, pGroup, True) Then
            Exit Sub
        End If

    End Sub

    Private Sub SyncWriteMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SyncWriteMI.Click

        Dim node As TreeNode = GroupsTV.SelectedNode

        If node Is Nothing Or Not TypeOf (node.Tag) Is OPCGroup Then
            Exit Sub
        End If

        Dim pGroup As OPCGroup = DirectCast(node.Tag, OPCGroup)

        If Not New WriteItemsDlg().ShowDialog(m_aServer, pGroup, False) Then
            Exit Sub
        End If

    End Sub

    Private Sub AsyncWriteMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AsyncWriteMI.Click

        Dim node As TreeNode = GroupsTV.SelectedNode

        If node Is Nothing Or Not TypeOf (node.Tag) Is OPCGroup Then
            Exit Sub
        End If

        Dim pGroup As OPCGroup = DirectCast(node.Tag, OPCGroup)

        If Not New WriteItemsDlg().ShowDialog(m_aServer, pGroup, True) Then
            Exit Sub
        End If

    End Sub

    Private Sub BrowseItemsMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BrowseItemsMI.Click

        Dim node As TreeNode = GroupsTV.SelectedNode

        If node Is Nothing Or Not TypeOf (node.Tag) Is OPCServer Then
            Exit Sub
        End If

        Dim pServer As OPCServer = DirectCast(node.Tag, OPCServer)

        If Not New BrowseItemsDlg().ShowDialog(pServer) Then
            Exit Sub
        End If

    End Sub

    Private Sub AsyncRefreshMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AsyncRefreshMI.Click

        Dim node As TreeNode = GroupsTV.SelectedNode

        If node Is Nothing Or Not TypeOf (node.Tag) Is OPCGroup Then
            Exit Sub
        End If

        Dim pGroup As OPCGroup = DirectCast(node.Tag, OPCGroup)

        Try
            Dim cancelID As Int32
            pGroup.AsyncRefresh(OPCDataSource.OPCDevice, m_aServer.NextHandle(), cancelID)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub RemoveGroupMI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RemoveGroupMI.Click

        Dim node As TreeNode = GroupsTV.SelectedNode

        If node Is Nothing Or Not TypeOf (node.Tag) Is OPCGroup Then
            Exit Sub
        End If

        Dim pGroup As OPCGroup = DirectCast(node.Tag, OPCGroup)

        pGroup.IsSubscribed = False
        pGroup.IsActive = False
        RaiseEvent GroupSubscribed(pGroup)

        Try
            m_aServer.RemoveGroup(pGroup.ClientHandle)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Dim pServer As OPCServer = pGroup.Parent
        Dim pGroups As OPCGroups = pServer.OPCGroups
        pGroups.Remove(pGroup.ServerHandle)

        node.Remove()

    End Sub

    Private Sub RemoveItemMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveItemMI.Click

        Dim node As TreeNode = GroupsTV.SelectedNode

        If node Is Nothing Or Not TypeOf (node.Tag) Is OPCItem Then
            Exit Sub
        End If

        Dim pItem As OPCItem = DirectCast(node.Tag, OPCItem)
        Dim pGroup As OPCGroup = pItem.Parent
        Dim pItems As OPCItems = pGroup.OPCItems

        Dim aGroup As Group = m_aServer.GetGroup(pGroup)
        aGroup.RemoveItem(pItem.ClientHandle)

        Try
            Dim serverHandles(1) As Int32
            Dim errors As Array

            serverHandles(1) = pItem.ServerHandle
            pItems.Remove(1, serverHandles, errors)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        node.Remove()

    End Sub

    Private Sub EditItemMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditItemMI.Click

        Dim node As TreeNode = GroupsTV.SelectedNode

        If node Is Nothing Or Not TypeOf (node.Tag) Is OPCItem Then
            Exit Sub
        End If

        Dim pItem As OPCItem = DirectCast(node.Tag, OPCItem)

        If Not New EditItemDlg().ShowDialog(pItem) Then
            Exit Sub
        End If

    End Sub
End Class
