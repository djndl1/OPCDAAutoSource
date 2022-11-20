'============================================================================
' TITLE: GroupUpdatesCtrl.vb
'
' CONTENTS:
' 
' A control used to display asynchronous updates from the server.
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

Public Class GroupUpdatesCtrl
    Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call    
        Static Resources As Resources = New Resources
        UpdatesLV.SmallImageList = Resources.ImageList

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
    Friend WithEvents UpdatesLV As System.Windows.Forms.ListView
    Friend WithEvents PopupMenu As System.Windows.Forms.ContextMenu
    Friend WithEvents ClearMI As System.Windows.Forms.MenuItem
    Friend WithEvents ViewMI As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.UpdatesLV = New System.Windows.Forms.ListView
        Me.PopupMenu = New System.Windows.Forms.ContextMenu
        Me.ClearMI = New System.Windows.Forms.MenuItem
        Me.ViewMI = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'UpdatesLV
        '
        Me.UpdatesLV.ContextMenu = Me.PopupMenu
        Me.UpdatesLV.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UpdatesLV.FullRowSelect = True
        Me.UpdatesLV.Location = New System.Drawing.Point(0, 0)
        Me.UpdatesLV.Name = "UpdatesLV"
        Me.UpdatesLV.Size = New System.Drawing.Size(416, 400)
        Me.UpdatesLV.TabIndex = 0
        Me.UpdatesLV.View = System.Windows.Forms.View.Details
        '
        'PopupMenu
        '
        Me.PopupMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.ClearMI, Me.MenuItem2, Me.ViewMI})
        '
        'ClearMI
        '
        Me.ClearMI.Index = 0
        Me.ClearMI.Text = "Clear"
        '
        'ViewMI
        '
        Me.ViewMI.Index = 2
        Me.ViewMI.Text = "View..."
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "-"
        '
        'GroupUpdatesCtrl
        '
        Me.Controls.Add(Me.UpdatesLV)
        Me.Name = "GroupUpdatesCtrl"
        Me.Size = New System.Drawing.Size(416, 400)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Class ItemHandle
        Public Group As OPCGroup
        Public Handle As Int32
        Public Value As Object
    End Class

    Private m_aServer As Server
    Private m_callbacks As Hashtable

    Delegate Sub OnDataChange_EventHandler(ByRef pGroup As OPCGroup, ByVal NumItems As Integer, ByRef ClientHandles As System.Array, ByRef ItemValues As System.Array, ByRef Qualities As System.Array, ByRef TimeStamps As System.Array)

    Class GroupCallback

        Public Shared Control As GroupUpdatesCtrl
        Public WithEvents Group As OPCGroup

        Private Sub Group_DataChange(ByVal TransactionID As Integer, ByVal NumItems As Integer, ByRef ClientHandles As System.Array, ByRef ItemValues As System.Array, ByRef Qualities As System.Array, ByRef TimeStamps As System.Array) Handles Group.DataChange
            Control.OnDataChange(Group, NumItems, ClientHandles, ItemValues, Qualities, TimeStamps)
        End Sub

        Private Sub Group_AsyncReadComplete(ByVal TransactionID As Integer, ByVal NumItems As Integer, ByRef ClientHandles As System.Array, ByRef ItemValues As System.Array, ByRef Qualities As System.Array, ByRef TimeStamps As System.Array, ByRef Errors As System.Array) Handles Group.AsyncReadComplete
            Control.OnReadComplete(Group, NumItems, ClientHandles, ItemValues, Qualities, TimeStamps, Errors)
        End Sub

        Private Sub Group_AsyncWriteComplete(ByVal TransactionID As Integer, ByVal NumItems As Integer, ByRef ClientHandles As System.Array, ByRef Errors As System.Array) Handles Group.AsyncWriteComplete
            Control.OnWriteComplete(Group, NumItems, ClientHandles, Errors)
        End Sub
    End Class

    Public Sub Initialize(ByRef server As Server)

        m_aServer = server
        m_callbacks = New Hashtable

        UpdatesLV.Clear()
        GroupCallback.Control = Me

        If m_aServer Is Nothing Then
            Exit Sub
        End If

        UpdatesLV.Columns.Add("Item ID", -2, HorizontalAlignment.Left)
        UpdatesLV.Columns.Add("Value", -2, HorizontalAlignment.Left)
        UpdatesLV.Columns.Add("Quality", -2, HorizontalAlignment.Left)
        UpdatesLV.Columns.Add("Timestamp", -2, HorizontalAlignment.Left)
        UpdatesLV.Columns.Add("Result", -2, HorizontalAlignment.Left)

        For Each column As ColumnHeader In UpdatesLV.Columns
            column.Width = -2
        Next

    End Sub

    Public Sub GroupSubscribed(ByRef pGroup As OPCGroup)

        If Not pGroup.IsSubscribed Then

            Dim items As ArrayList = New ArrayList

            For Each listItem As ListViewItem In UpdatesLV.Items

                Dim handle As ItemHandle = DirectCast(listItem.Tag, ItemHandle)

                If handle.Group.ClientHandle = pGroup.ClientHandle Then
                    items.Add(listItem)
                End If

            Next

            For Each listItem As ListViewItem In items
                listItem.Remove()
            Next

            m_callbacks.Remove(pGroup.ClientHandle)

        Else

            If Not m_callbacks.ContainsKey(pGroup.ClientHandle) Then
                Dim callback As GroupCallback = New GroupCallback
                callback.Group = pGroup
                m_callbacks.Add(pGroup.ClientHandle, callback)
            End If

        End If

    End Sub

    Public Sub OnDataChange(ByRef pGroup As OPCGroup, ByVal NumItems As Integer, ByRef ClientHandles As System.Array, ByRef ItemValues As System.Array, ByRef Qualities As System.Array, ByRef TimeStamps As System.Array)

        If InvokeRequired Then

            Dim args(6) As Object

            args(0) = pGroup
            args(1) = NumItems
            args(2) = ClientHandles
            args(3) = ItemValues
            args(4) = Qualities
            args(5) = TimeStamps

            Invoke(New OnDataChange_EventHandler(AddressOf OnDataChange), args)

        End If

        For Each listItem As ListViewItem In UpdatesLV.Items
            listItem.ForeColor = Color.Gray
        Next

        For ii As Int32 = 1 To NumItems

            Try
                UpdateItem(pGroup, ClientHandles(ii), ItemValues(ii), Qualities(ii), TimeStamps(ii))
            Catch ex As Exception
                'MsgBox(ex.Message)
            End Try

        Next

        For Each column As ColumnHeader In UpdatesLV.Columns
            column.Width = -2
        Next

    End Sub
    Public Sub OnReadComplete(ByRef pGroup As OPCGroup, ByVal NumItems As Integer, ByRef ClientHandles As System.Array, ByRef ItemValues As System.Array, ByRef Qualities As System.Array, ByRef TimeStamps As System.Array, ByRef Errors As Array)

        If InvokeRequired Then

            Dim args(7) As Object

            args(0) = pGroup
            args(1) = NumItems
            args(2) = ClientHandles
            args(3) = ItemValues
            args(4) = Qualities
            args(5) = TimeStamps
            args(6) = Errors

            Invoke(New OnDataChange_EventHandler(AddressOf OnDataChange), args)

        End If

        For ii As Int32 = 1 To NumItems

            Try
                AddItemRead(pGroup, ClientHandles(ii), ItemValues(ii), Qualities(ii), TimeStamps(ii), Errors(ii))
            Catch ex As Exception
                'MsgBox(ex.Message)
            End Try

        Next

        For Each column As ColumnHeader In UpdatesLV.Columns
            column.Width = -2
        Next

    End Sub

    Public Sub OnWriteComplete(ByRef pGroup As OPCGroup, ByVal NumItems As Integer, ByRef ClientHandles As System.Array, ByRef Errors As Array)

        If InvokeRequired Then

            Dim args(4) As Object

            args(0) = pGroup
            args(1) = NumItems
            args(2) = ClientHandles
            args(3) = Errors

            Invoke(New OnDataChange_EventHandler(AddressOf OnDataChange), args)

        End If

        For ii As Int32 = 1 To NumItems

            Try
                AddItemWrite(pGroup, ClientHandles(ii), Errors(ii))
            Catch ex As Exception
                'MsgBox(ex.Message)
            End Try

        Next

        For Each column As ColumnHeader In UpdatesLV.Columns
            column.Width = -2
        Next

    End Sub

    Private Sub UpdateItem(ByRef pGroup As OPCGroup, ByVal clientHandle As Int32, ByVal value As Object, ByVal quality As Int16, ByVal timestamp As Date)

        For Each listItem As ListViewItem In UpdatesLV.Items

            Dim handle As ItemHandle = DirectCast(listItem.Tag, ItemHandle)

            If (handle.Group Is pGroup And handle.Handle = clientHandle) Then

                listItem.ForeColor = Color.Empty

                listItem.SubItems(1).Text = Opc.Convert.ToString(value)
                listItem.SubItems(2).Text = Opc.Convert.ToString(quality)
                listItem.SubItems(3).Text = Opc.Convert.ToString(CType(timestamp, DateTime))
                listItem.SubItems(4).Text = ""

                handle.Value = value

                Exit Sub
            End If
        Next

        AddItemUpdate(pGroup, clientHandle, value, quality, timestamp)

    End Sub

    Private Sub AddItemUpdate(ByRef pGroup As OPCGroup, ByVal clientHandle As Int32, ByVal value As Object, ByVal quality As Int16, ByVal timestamp As Date)

        Dim group As Group = m_aServer.GetGroup(pGroup)

        If group Is Nothing Then
            Exit Sub
        End If

        Dim item As Item = group.Items.Item(clientHandle)

        If item Is Nothing Then
            Exit Sub
        End If

        Dim handle As ItemHandle = New ItemHandle
        handle.Group = pGroup
        handle.Handle = clientHandle

        Dim listItem As ListViewItem = New ListViewItem(item.ItemID)
        listItem.ImageIndex = Resources.ITEM
        listItem.Tag = handle

        listItem.SubItems.Add(Opc.Convert.ToString(value))
        listItem.SubItems.Add(Opc.Convert.ToString(quality))
        listItem.SubItems.Add(Opc.Convert.ToString(CType(timestamp, DateTime)))
        listItem.SubItems.Add("")

        handle.Value = value

        UpdatesLV.Items.Add(listItem)

    End Sub

    Private Sub AddItemRead(ByRef pGroup As OPCGroup, ByVal clientHandle As Int32, ByVal value As Object, ByVal quality As Int16, ByVal timestamp As Date, ByVal result As Int32)

        Dim group As Group = m_aServer.GetGroup(pGroup)

        If group Is Nothing Then
            Exit Sub
        End If

        Dim item As Item = group.Items.Item(clientHandle)

        If item Is Nothing Then
            Exit Sub
        End If

        Dim handle As ItemHandle = New ItemHandle
        handle.Group = pGroup
        handle.Handle = clientHandle

        Dim listItem As ListViewItem = New ListViewItem(item.ItemID)
        listItem.ImageIndex = Resources.ALTERNATE_ITEM
        listItem.Tag = handle
        listItem.ForeColor = Color.Firebrick

        listItem.SubItems.Add(Opc.Convert.ToString(value))
        listItem.SubItems.Add(Opc.Convert.ToString(quality))
        listItem.SubItems.Add(Opc.Convert.ToString(CType(timestamp, DateTime)))
        listItem.SubItems.Add(String.Format("&H{0:X8}", result))

        handle.Value = value

        UpdatesLV.Items.Add(listItem)

    End Sub

    Private Sub AddItemWrite(ByRef pGroup As OPCGroup, ByVal clientHandle As Int32, ByVal result As Int32)

        Dim group As Group = m_aServer.GetGroup(pGroup)

        If group Is Nothing Then
            Exit Sub
        End If

        Dim item As Item = group.Items.Item(clientHandle)

        If item Is Nothing Then
            Exit Sub
        End If

        Dim handle As ItemHandle = New ItemHandle
        handle.Group = pGroup
        handle.Handle = clientHandle

        Dim listItem As ListViewItem = New ListViewItem(item.ItemID)
        listItem.ImageIndex = Resources.ALTERNATE_ITEM
        listItem.Tag = handle
        listItem.ForeColor = Color.ForestGreen

        listItem.SubItems.Add("")
        listItem.SubItems.Add("")
        listItem.SubItems.Add("")
        listItem.SubItems.Add(String.Format("&H{0:X8}", result))

        UpdatesLV.Items.Add(listItem)

    End Sub

    Private Sub ClearMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearMI.Click
        UpdatesLV.Items.Clear()
    End Sub

    Private Sub ViewMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewMI.Click

        If UpdatesLV.SelectedItems.Count <> 1 Then
            Exit Sub
        End If

        Dim handle As ItemHandle = DirectCast(UpdatesLV.SelectedItems(0).Tag, ItemHandle)

        If Not handle.Value Is Nothing Then

            If handle.Value.GetType().IsArray Then
                handle.Value = New EditArrayDlg().ShowDialog(handle.Value, True)
            End If

        End If

    End Sub
End Class
