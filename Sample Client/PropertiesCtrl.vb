'============================================================================
' TITLE: PropertiesCtrl.vb
'
' CONTENTS:
' 
' A control used to display the properties of an item.
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

Public Class PropertiesCtrl
    Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Static Resources As Resources = New Resources
        PropertiesLV.SmallImageList = Resources.ImageList

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
    Friend WithEvents PropertiesLV As System.Windows.Forms.ListView
    Friend WithEvents PopupMenu As System.Windows.Forms.ContextMenu
    Friend WithEvents ViewMI As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.PropertiesLV = New System.Windows.Forms.ListView
        Me.PopupMenu = New System.Windows.Forms.ContextMenu
        Me.ViewMI = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'PropertiesLV
        '
        Me.PropertiesLV.ContextMenu = Me.PopupMenu
        Me.PropertiesLV.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PropertiesLV.FullRowSelect = True
        Me.PropertiesLV.Location = New System.Drawing.Point(0, 0)
        Me.PropertiesLV.Name = "PropertiesLV"
        Me.PropertiesLV.Size = New System.Drawing.Size(376, 384)
        Me.PropertiesLV.TabIndex = 0
        Me.PropertiesLV.View = System.Windows.Forms.View.Details
        '
        'PopupMenu
        '
        Me.PopupMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.ViewMI})
        '
        'ViewMI
        '
        Me.ViewMI.Index = 0
        Me.ViewMI.Text = "View..."
        '
        'PropertiesCtrl
        '
        Me.Controls.Add(Me.PropertiesLV)
        Me.Name = "PropertiesCtrl"
        Me.Size = New System.Drawing.Size(376, 384)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private m_pServer As OPCServer

    Public Sub Initialize(ByRef pServer As OPCServer)

        PropertiesLV.Clear()
        PropertiesLV.Columns.Add("ID", -2, HorizontalAlignment.Left)
        PropertiesLV.Columns.Add("Description", -2, HorizontalAlignment.Left)
        PropertiesLV.Columns.Add("Data Type", -2, HorizontalAlignment.Left)
        PropertiesLV.Columns.Add("Value", -2, HorizontalAlignment.Left)
        PropertiesLV.Columns.Add("Error", -2, HorizontalAlignment.Left)
        PropertiesLV.Columns.Add("Item ID", -2, HorizontalAlignment.Left)

        m_pServer = pServer

    End Sub

    Public Sub ShowProperties(ByVal itemID As String)

        PropertiesLV.Items.Clear()

        If itemID Is Nothing Then
            Exit Sub
        End If

        Try

            Dim count As Int32
            Dim propertyIDs As Array
            Dim descriptions As Array
            Dim dataTypes As Array

            m_pServer.QueryAvailableProperties(itemID, count, propertyIDs, descriptions, dataTypes)

            Dim values As Array
            Dim errors As Array

            Try
                m_pServer.GetItemProperties(itemID, count, propertyIDs, values, errors)
            Catch ex As Exception
            End Try

            Dim itemIDs As Array
            Dim itemIDErrors As Array

            Try
                m_pServer.LookupItemIDs(itemID, count, propertyIDs, itemIDs, itemIDErrors)
            Catch ex As Exception
            End Try

            For ii As Int32 = 1 To count

                Dim item As ListViewItem = New ListViewItem(CStr(propertyIDs(ii)))

                item.ImageIndex = Resources.ITEM_PROPERTY
                item.Tag = values(ii)

                item.SubItems.Add(descriptions(ii))
                item.SubItems.Add(Opc.Convert.ToString(Opc.Type.GetType(dataTypes(ii))))

                If errors(ii) >= 0 Then
                    item.SubItems.Add(Opc.Convert.ToString(values(ii)))
                    item.SubItems.Add("")
                Else
                    item.SubItems.Add("")
                    item.SubItems.Add(String.Format("&H{0:X8}", errors(ii)))
                End If

                If itemIDErrors(ii) >= 0 Then
                    item.SubItems.Add(itemIDs(ii))
                Else
                    item.SubItems.Add("")
                End If

                PropertiesLV.Items.Add(item)

            Next

            For Each column As ColumnHeader In PropertiesLV.Columns
                column.Width = -2

            Next

        Catch ex As Exception
        End Try

    End Sub

    Private Sub ViewMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewMI.Click

        If PropertiesLV.SelectedItems.Count <> 1 Then
            Exit Sub
        End If

        Dim value As Object = PropertiesLV.SelectedItems(0).Tag

        If Not value Is Nothing Then

            If value.GetType().IsArray Then
                value = New EditArrayDlg().ShowDialog(value, True)
            End If

        End If

    End Sub

    Private Sub PropertiesLV_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PropertiesLV.MouseDown

        If e.Button <> MouseButtons.Right Then
            Exit Sub
        End If

        ViewMI.Enabled = False

        Dim clickedItem As ListViewItem = PropertiesLV.GetItemAt(e.X, e.Y)

        If clickedItem Is Nothing Then
            Exit Sub
        End If

        If PropertiesLV.SelectedItems.Count = 0 Then
            clickedItem.Selected = True
        End If

        If PropertiesLV.SelectedItems.Count = 1 Then
            ViewMI.Enabled = True
        End If

    End Sub

    Private Sub PropertiesLV_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles PropertiesLV.DoubleClick
        ViewMI_Click(sender, e)
    End Sub
End Class
