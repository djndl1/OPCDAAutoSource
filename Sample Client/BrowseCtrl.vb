'============================================================================
' TITLE: BrowseCtrl.vb
'
' CONTENTS:
' 
' A control used to browse the address space of a server.
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
Imports System.Collections


Public Class BrowseCtrl

    Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Static Resources As Resources = New Resources
        BrowseTV.ImageList = Resources.ImageList

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
    Friend WithEvents BrowseTV As System.Windows.Forms.TreeView
    Friend WithEvents PopupMenu As System.Windows.Forms.ContextMenu
    Friend WithEvents SetFiltersMI As System.Windows.Forms.MenuItem
    Friend WithEvents RefreshMI As System.Windows.Forms.MenuItem
    Friend WithEvents SelectMI As System.Windows.Forms.MenuItem
    Friend WithEvents SelectChildrenMI As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.BrowseTV = New System.Windows.Forms.TreeView
        Me.PopupMenu = New System.Windows.Forms.ContextMenu
        Me.SetFiltersMI = New System.Windows.Forms.MenuItem
        Me.RefreshMI = New System.Windows.Forms.MenuItem
        Me.SelectMI = New System.Windows.Forms.MenuItem
        Me.SelectChildrenMI = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'BrowseTV
        '
        Me.BrowseTV.ContextMenu = Me.PopupMenu
        Me.BrowseTV.Dock = System.Windows.Forms.DockStyle.Fill
        Me.BrowseTV.ImageIndex = -1
        Me.BrowseTV.Location = New System.Drawing.Point(0, 0)
        Me.BrowseTV.Name = "BrowseTV"
        Me.BrowseTV.SelectedImageIndex = -1
        Me.BrowseTV.Size = New System.Drawing.Size(408, 400)
        Me.BrowseTV.TabIndex = 0
        '
        'PopupMenu
        '
        Me.PopupMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.SelectMI, Me.SelectChildrenMI, Me.MenuItem3, Me.SetFiltersMI, Me.RefreshMI})
        '
        'SetFiltersMI
        '
        Me.SetFiltersMI.Index = 3
        Me.SetFiltersMI.Text = "Set Filters..."
        '
        'RefreshMI
        '
        Me.RefreshMI.Index = 4
        Me.RefreshMI.Text = "Refresh"
        '
        'SelectMI
        '
        Me.SelectMI.Index = 0
        Me.SelectMI.Text = "Select"
        '
        'SelectChildrenMI
        '
        Me.SelectChildrenMI.Index = 1
        Me.SelectChildrenMI.Text = "Select Children"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 2
        Me.MenuItem3.Text = "-"
        '
        'BrowseCtrl
        '
        Me.Controls.Add(Me.BrowseTV)
        Me.Name = "BrowseCtrl"
        Me.Size = New System.Drawing.Size(408, 400)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private browser As OPCBrowser

    Public Delegate Sub ItemPicked_EventHandler(ByVal itemID As String)
    Public Delegate Sub ItemSelected_EventHandler(ByVal itemID As String)

    Public Event ItemPicked As ItemPicked_EventHandler
    Public Event ItemSelected As ItemSelected_EventHandler

    Public Sub Connect(ByVal pServer As OPCServer)

        BrowseTV.Nodes.Clear()

        browser = pServer.CreateBrowser()

        Browse(Nothing)

    End Sub
    Private Sub GetNames(ByRef node As TreeNode, ByRef names As ArrayList)

        If node Is Nothing Then
            Exit Sub
        End If

        If Not node.Parent Is Nothing Then
            GetNames(node.Parent, names)
        End If

        names.Add(node.Text)

    End Sub
    Private Sub Browse(ByRef parent As TreeNode)

        Dim nodes As TreeNodeCollection

        If parent Is Nothing Then
            browser.MoveToRoot()
            nodes = BrowseTV.Nodes
        Else
            Dim names As ArrayList = New ArrayList

            GetNames(parent, names)
            browser.MoveTo(names.ToArray(GetType(String)))

            nodes = parent.Nodes
        End If

        nodes.Clear()

        browser.ShowBranches()

        For Each branch As String In browser
            Dim node As TreeNode

            node = New TreeNode(branch)
            node.ImageIndex = Resources.FOLDER
            node.SelectedImageIndex = Resources.OPEN_FOLDER

            Try
                node.Tag = browser.GetItemID(branch)
            Catch ex As Exception
                node.Tag = Nothing
            End Try

            node.Nodes.Add(New TreeNode)

            nodes.Add(node)

        Next

        browser.ShowLeafs()

        For Each leaf As String In browser
            Dim node As TreeNode

            node = New TreeNode(leaf)
            node.ImageIndex = Resources.ITEM
            node.SelectedImageIndex = Resources.ITEM

            Try
                node.Tag = browser.GetItemID(leaf)
            Catch ex As Exception
                node.Tag = Nothing
            End Try

            nodes.Add(node)
        Next

    End Sub
    Private Sub BrowseTV_BeforeExpand(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles BrowseTV.BeforeExpand

        If e.Node.Nodes.Count = 1 Then
            Dim node As TreeNode = e.Node.Nodes(0)

            If node.Text = "" Then
                Browse(e.Node)
            End If
        End If

    End Sub

    Private Sub SetFiltersMI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SetFiltersMI.Click

        Dim node As TreeNode = BrowseTV.SelectedNode

        If node Is Nothing Then
            Exit Sub
        End If

        Dim dialog As BrowseFiltersDlg = New BrowseFiltersDlg

        If Not dialog.ShowDialog(browser) Then
            Exit Sub
        End If

        Browse(node)

    End Sub

    Private Sub BrowseTV_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BrowseTV.MouseDown

        If e.Button <> MouseButtons.Right Then
            Exit Sub
        End If

        SelectMI.Enabled = False
        SelectChildrenMI.Enabled = False
        SetFiltersMI.Enabled = False
        RefreshMI.Enabled = False

        Dim clickedNode As TreeNode = BrowseTV.GetNodeAt(e.X, e.Y)

        If clickedNode Is Nothing Then
            Exit Sub
        End If

        BrowseTV.SelectedNode = clickedNode

        If TypeOf (clickedNode.Tag) Is String Then
            SelectMI.Enabled = True
        End If

        If clickedNode.Nodes.Count > 0 Then
            SelectChildrenMI.Enabled = True
        End If

        SetFiltersMI.Enabled = True
        RefreshMI.Enabled = True

    End Sub

    Private Sub SelectChildrenMI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SelectChildrenMI.Click

        Dim node As TreeNode = BrowseTV.SelectedNode

        If node Is Nothing Then
            Exit Sub
        End If

        For Each child As TreeNode In node.Nodes
            If TypeOf (child.Tag) Is String Then
                RaiseEvent ItemPicked(DirectCast(child.Tag, String))
            End If
        Next

    End Sub

    Private Sub SelectMI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SelectMI.Click

        Dim node As TreeNode = BrowseTV.SelectedNode

        If node Is Nothing Then
            Exit Sub
        End If

        If TypeOf (node.Tag) Is String Then
            RaiseEvent ItemPicked(DirectCast(node.Tag, String))
        End If

    End Sub

    Private Sub BrowseTV_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles BrowseTV.AfterSelect

        Dim node As TreeNode = BrowseTV.SelectedNode

        If node Is Nothing Then
            Exit Sub
        End If

        If TypeOf (node.Tag) Is String Then
            RaiseEvent ItemSelected(DirectCast(node.Tag, String))
        End If

    End Sub

    Private Sub BrowseTV_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BrowseTV.DoubleClick
        SelectMI_Click(sender, e)
    End Sub
End Class
