'============================================================================
' TITLE: ItemsCtrl.vb
'
' CONTENTS:
' 
' A control used to manage a set of items used for an operation.
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
Imports Opc.Da.SampleClient
Imports System.Collections

Public Class ItemsCtrl
    Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Static Resources As Resources = New Resources
        ItemsLV.SmallImageList = Resources.ImageList

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
    Friend WithEvents ItemsLV As System.Windows.Forms.ListView
    Friend WithEvents PopupMenu As System.Windows.Forms.ContextMenu
    Friend WithEvents EditMI As System.Windows.Forms.MenuItem
    Friend WithEvents RemoveMI As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents ViewMI As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ItemsLV = New System.Windows.Forms.ListView
        Me.PopupMenu = New System.Windows.Forms.ContextMenu
        Me.EditMI = New System.Windows.Forms.MenuItem
        Me.RemoveMI = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.ViewMI = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'ItemsLV
        '
        Me.ItemsLV.ContextMenu = Me.PopupMenu
        Me.ItemsLV.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ItemsLV.FullRowSelect = True
        Me.ItemsLV.Location = New System.Drawing.Point(0, 0)
        Me.ItemsLV.Name = "ItemsLV"
        Me.ItemsLV.Size = New System.Drawing.Size(432, 328)
        Me.ItemsLV.TabIndex = 1
        Me.ItemsLV.View = System.Windows.Forms.View.Details
        '
        'PopupMenu
        '
        Me.PopupMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.EditMI, Me.RemoveMI, Me.MenuItem1, Me.ViewMI})
        '
        'EditMI
        '
        Me.EditMI.Index = 0
        Me.EditMI.Text = "Edit..."
        '
        'RemoveMI
        '
        Me.RemoveMI.Index = 1
        Me.RemoveMI.Text = "Remove"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 2
        Me.MenuItem1.Text = "-"
        '
        'ViewMI
        '
        Me.ViewMI.Index = 3
        Me.ViewMI.Text = "View..."
        '
        'ItemsCtrl
        '
        Me.Controls.Add(Me.ItemsLV)
        Me.Name = "ItemsCtrl"
        Me.Size = New System.Drawing.Size(432, 328)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Enum Mode
        Read = 1
        Write = 2
        Subscribe = 3
    End Enum

    Private m_pItems As OPCItems
    Private m_mode As Mode

    Public Sub Initialize(ByRef pItems As OPCItems, ByVal mode As Mode)

        If pItems Is Nothing Then
            Exit Sub
        End If

        m_pItems = pItems
        m_mode = mode

        ItemsLV.Clear()
        ItemsLV.Columns.Add("Item ID", -2, HorizontalAlignment.Left)

        If mode = mode.Subscribe Then

            ItemsLV.Columns.Add("Is Active", -2, HorizontalAlignment.Left)
            ItemsLV.Columns.Add("Req Type", -2, HorizontalAlignment.Left)
            ItemsLV.Columns.Add("Access Path", -2, HorizontalAlignment.Left)

            Dim listItem As ListViewItem = New ListViewItem
            listItem.ImageIndex = Resources.ALTERNATE_ITEM
            ItemsLV.Items.Add(listItem)

            Dim item As Item = New Item

            item.ItemID = "<default>"
            item.Active = m_pItems.DefaultIsActive
            item.ReqType = Opc.Type.GetType(m_pItems.DefaultRequestedDataType)
            item.AccessPath = m_pItems.DefaultAccessPath

            UpdateItem(item, listItem, False)

        ElseIf mode = mode.Read Then
            'nothing to do.

        ElseIf mode = mode.Write Then

            ItemsLV.Columns.Add("Value", -2, HorizontalAlignment.Left)
            ItemsLV.Columns.Add("Quality", -2, HorizontalAlignment.Left)
            ItemsLV.Columns.Add("Timestamp", -2, HorizontalAlignment.Left)

        End If

        For Each column As ColumnHeader In ItemsLV.Columns
            column.Width = -2
        Next

    End Sub

    Public Sub Initialize(ByRef items() As Item, ByRef errors As Array, ByVal mode As Mode)

        m_pItems = Nothing
        m_mode = mode

        ItemsLV.Clear()
        ItemsLV.Columns.Add("Item ID", -2, HorizontalAlignment.Left)

        If mode = mode.Subscribe Then

            ItemsLV.Columns.Add("Is Active", -2, HorizontalAlignment.Left)
            ItemsLV.Columns.Add("Req Type", -2, HorizontalAlignment.Left)
            ItemsLV.Columns.Add("Access Path", -2, HorizontalAlignment.Left)
            ItemsLV.Columns.Add("Result", -2, HorizontalAlignment.Left)

        ElseIf mode = mode.Read Then

            ItemsLV.Columns.Add("Value", -2, HorizontalAlignment.Left)
            ItemsLV.Columns.Add("Quality", -2, HorizontalAlignment.Left)
            ItemsLV.Columns.Add("Timestamp", -2, HorizontalAlignment.Left)
            ItemsLV.Columns.Add("Result", -2, HorizontalAlignment.Left)

        ElseIf mode = mode.Write Then

            ItemsLV.Columns.Add("Value", -2, HorizontalAlignment.Left)
            ItemsLV.Columns.Add("Quality", -2, HorizontalAlignment.Left)
            ItemsLV.Columns.Add("Timestamp", -2, HorizontalAlignment.Left)
            ItemsLV.Columns.Add("Result", -2, HorizontalAlignment.Left)

        End If

        For ii As Int32 = 1 To items.Length - 1

            Dim listItem As ListViewItem = New ListViewItem
            listItem.ImageIndex = Resources.ALTERNATE_ITEM
            listItem.Tag = items(ii)
            ItemsLV.Items.Add(listItem)

            UpdateItem(items(ii), listItem, True)

            listItem.SubItems(4).Text = String.Format("&H{0:X8}", errors(ii))

        Next

        For Each column As ColumnHeader In ItemsLV.Columns
            column.Width = -2
        Next

    End Sub

    Public Function GetItems() As Item()
        Dim items As ArrayList = New ArrayList

        For Each current As ListViewItem In ItemsLV.Items
            items.Add(current.Tag)
        Next

        GetItems = DirectCast(items.ToArray(GetType(Item)), Item())
    End Function

    Public Sub ClearItems()
        ItemsLV.Items.Clear()
    End Sub

    Public Sub AddItem(ByVal itemID As String)

        If m_pItems Is Nothing Then
            Exit Sub
        End If

        Dim item As Item = New Item
        item.ItemID = itemID
        item.Active = m_pItems.DefaultIsActive

        Dim listItem As ListViewItem = New ListViewItem

        listItem.Text = itemID
        listItem.ImageIndex = Resources.ITEM
        listItem.Tag = item

        ItemsLV.Items.Add(listItem)

        UpdateItem(item, listItem, False)

        For Each column As ColumnHeader In ItemsLV.Columns
            column.Width = -2
        Next

    End Sub
    Public Sub AddItem(ByRef pItem As OPCItem)

        If m_pItems Is Nothing Then
            Exit Sub
        End If

        Dim aItem As Item = Item.NewItem(pItem)

        Dim listItem As ListViewItem = New ListViewItem

        listItem.Text = aItem.ItemID
        listItem.ImageIndex = Resources.ITEM
        listItem.Tag = aItem

        ItemsLV.Items.Add(listItem)

        UpdateItem(aItem, listItem, False)

        For Each column As ColumnHeader In ItemsLV.Columns
            column.Width = -2
        Next

    End Sub
    Private Sub UpdateItem(ByRef item As Item, ByRef listItem As ListViewItem, ByVal isResult As Boolean)

        listItem.Tag = item

        While listItem.SubItems.Count < ItemsLV.Columns.Count
            listItem.SubItems.Add("")
        End While

        If m_mode = Mode.Subscribe Then

            listItem.SubItems(0).Text = Opc.Convert.ToString(item.ItemID)
            listItem.SubItems(1).Text = Opc.Convert.ToString(item.Active)
            listItem.SubItems(2).Text = Opc.Convert.ToString(item.ReqType)
            listItem.SubItems(3).Text = Opc.Convert.ToString(item.AccessPath)

        ElseIf m_mode = Mode.Read Then

            listItem.SubItems(0).Text = Opc.Convert.ToString(item.ItemID)

            If isResult Then

                listItem.SubItems(1).Text = Opc.Convert.ToString(item.Value)
                listItem.SubItems(2).Text = Opc.Convert.ToString(item.Quality)
                listItem.SubItems(3).Text = Opc.Convert.ToString(item.Timestamp)

            End If

        ElseIf m_mode = Mode.Write Then

            listItem.SubItems(0).Text = Opc.Convert.ToString(item.ItemID)
            listItem.SubItems(1).Text = Opc.Convert.ToString(item.Value)
            listItem.SubItems(2).Text = Opc.Convert.ToString(item.Quality)
            listItem.SubItems(3).Text = Opc.Convert.ToString(item.Timestamp)

        End If

    End Sub
    Private Sub EditMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditMI.Click

        If ItemsLV.SelectedItems.Count <> 1 Then
            Exit Sub
        End If

        Dim selectedItem As ListViewItem = ItemsLV.SelectedItems(0)

        If Not TypeOf (selectedItem.Tag) Is Item Then
            Exit Sub
        End If

        Dim item As Item = DirectCast(selectedItem.Tag, Item)

        Dim dialog As EditItemDlg = New EditItemDlg

        If Not dialog.ShowDialog(item, m_mode) Then
            Exit Sub
        End If

        UpdateItem(item, selectedItem, False)

        For Each column As ColumnHeader In ItemsLV.Columns
            column.Width = -2
        Next

    End Sub

    Private Sub ItemsLV_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ItemsLV.MouseDown

        If e.Button <> MouseButtons.Right Then
            Exit Sub
        End If

        EditMI.Enabled = False
        RemoveMI.Enabled = False
        ViewMI.Enabled = False

        Dim clickedItem As ListViewItem = ItemsLV.GetItemAt(e.X, e.Y)

        If clickedItem Is Nothing Then
            Exit Sub
        End If

        If ItemsLV.SelectedItems.Count = 0 Then
            clickedItem.Selected = True
        End If

        If Not m_pItems Is Nothing Then
            If ItemsLV.SelectedItems.Count = 1 Then
                EditMI.Enabled = True
            End If

            RemoveMI.Enabled = True
        Else
            ViewMI.Enabled = True
        End If

    End Sub

    Private Sub RemoveMI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RemoveMI.Click

        Dim items As ArrayList = New ArrayList

        For Each listItem As ListViewItem In ItemsLV.SelectedItems
            items.Add(listItem)
        Next

        For Each listItem As ListViewItem In items
            listItem.Remove()
        Next

    End Sub

    Private Sub ViewMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewMI.Click

        If ItemsLV.SelectedItems.Count <> 1 Then
            Exit Sub
        End If

        Dim item As Item = DirectCast(ItemsLV.SelectedItems(0).Tag, Item)

        If Not item.Value Is Nothing Then

            If item.Value.GetType().IsArray Then
                item.Value = New EditArrayDlg().ShowDialog(item.Value, True)
            End If

        End If

    End Sub
End Class

