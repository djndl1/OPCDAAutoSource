'============================================================================
' TITLE: GroupItemsCtrl.vb
'
' CONTENTS:
' 
' A control used to select a set of items belonging to a group.
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

Public Class GroupItemsCtrl
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
    Friend WithEvents SelectMI As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ItemsLV = New System.Windows.Forms.ListView
        Me.PopupMenu = New System.Windows.Forms.ContextMenu
        Me.SelectMI = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'ItemsLV
        '
        Me.ItemsLV.ContextMenu = Me.PopupMenu
        Me.ItemsLV.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ItemsLV.FullRowSelect = True
        Me.ItemsLV.Location = New System.Drawing.Point(0, 0)
        Me.ItemsLV.Name = "ItemsLV"
        Me.ItemsLV.Size = New System.Drawing.Size(408, 392)
        Me.ItemsLV.TabIndex = 0
        Me.ItemsLV.View = System.Windows.Forms.View.Details
        '
        'PopupMenu
        '
        Me.PopupMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.SelectMI})
        '
        'SelectMI
        '
        Me.SelectMI.Index = 0
        Me.SelectMI.Text = "Select"
        '
        'GroupItemsCtrl
        '
        Me.Controls.Add(Me.ItemsLV)
        Me.Name = "GroupItemsCtrl"
        Me.Size = New System.Drawing.Size(408, 392)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Delegate Sub ItemPicked_EventHandler(ByRef item As OPCItem)

    Public Event ItemPicked As ItemPicked_EventHandler

    Public Sub Initialize(ByRef pItems As OPCItems)

        ItemsLV.Clear()
        ItemsLV.Columns.Add("Item ID", -2, HorizontalAlignment.Left)
        ItemsLV.Columns.Add("Data Type", -2, HorizontalAlignment.Left)
        ItemsLV.Columns.Add("Access Rights", -2, HorizontalAlignment.Left)

        For Each pItem As OPCItem In pItems

            Dim listitem As ListViewItem = New ListViewItem(pItem.ItemID)

            listitem.ImageIndex = Resources.ITEM
            listitem.Tag = pItem

            Dim accessRights As Object = [Enum].ToObject(GetType(AccessRights), pItem.AccessRights)

            listitem.SubItems.Add(Opc.Convert.ToString(Opc.Type.GetType(pItem.CanonicalDataType)))
            listitem.SubItems.Add(accessRights.ToString())

            ItemsLV.Items.Add(listitem)

        Next

    End Sub

    Private Sub SelectMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectMI.Click

        For Each listitem As ListViewItem In ItemsLV.SelectedItems
            RaiseEvent ItemPicked(listitem.Tag)
        Next

    End Sub

    Private Sub ItemsLV_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ItemsLV.DoubleClick
        SelectMI_Click(sender, e)
    End Sub

    Private Sub ItemsLV_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ItemsLV.MouseDown

        If e.Button <> MouseButtons.Right Then
            Exit Sub
        End If

        SelectMI.Enabled = False

        Dim clickedItem As ListViewItem = ItemsLV.GetItemAt(e.X, e.Y)

        If clickedItem Is Nothing Then
            Exit Sub
        End If

        If ItemsLV.SelectedItems.Count = 0 Then
            clickedItem.Selected = True
        End If

        SelectMI.Enabled = True

    End Sub

End Class
