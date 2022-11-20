'============================================================================
' TITLE: AddItemsDlg.vb
'
' CONTENTS:
' 
' A dialog used to add items to a group.
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

Public Class AddItemsDlg
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
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
    Friend WithEvents BottomPN As System.Windows.Forms.Panel
    Friend WithEvents LeftPN As System.Windows.Forms.Panel
    Friend WithEvents SplitterV As System.Windows.Forms.Splitter
    Friend WithEvents RightPN As System.Windows.Forms.Panel
    Friend WithEvents BrowseCTRL As VbNetSampleClient.BrowseCtrl
    Friend WithEvents ItemsCTRL As VbNetSampleClient.ItemsCtrl
    Friend WithEvents DoneBTN As System.Windows.Forms.Button
    Friend WithEvents NextBTN As System.Windows.Forms.Button
    Friend WithEvents BackBTN As System.Windows.Forms.Button
    Friend WithEvents CancelBTN As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.BottomPN = New System.Windows.Forms.Panel
        Me.BackBTN = New System.Windows.Forms.Button
        Me.NextBTN = New System.Windows.Forms.Button
        Me.DoneBTN = New System.Windows.Forms.Button
        Me.CancelBTN = New System.Windows.Forms.Button
        Me.BrowseCTRL = New VbNetSampleClient.BrowseCtrl
        Me.LeftPN = New System.Windows.Forms.Panel
        Me.SplitterV = New System.Windows.Forms.Splitter
        Me.RightPN = New System.Windows.Forms.Panel
        Me.ItemsCTRL = New VbNetSampleClient.ItemsCtrl
        Me.BottomPN.SuspendLayout()
        Me.LeftPN.SuspendLayout()
        Me.RightPN.SuspendLayout()
        Me.SuspendLayout()
        '
        'BottomPN
        '
        Me.BottomPN.Controls.Add(Me.BackBTN)
        Me.BottomPN.Controls.Add(Me.NextBTN)
        Me.BottomPN.Controls.Add(Me.DoneBTN)
        Me.BottomPN.Controls.Add(Me.CancelBTN)
        Me.BottomPN.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.BottomPN.Location = New System.Drawing.Point(0, 334)
        Me.BottomPN.Name = "BottomPN"
        Me.BottomPN.Size = New System.Drawing.Size(1016, 32)
        Me.BottomPN.TabIndex = 2
        '
        'BackBTN
        '
        Me.BackBTN.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BackBTN.Location = New System.Drawing.Point(858, 4)
        Me.BackBTN.Name = "BackBTN"
        Me.BackBTN.TabIndex = 3
        Me.BackBTN.Text = "< Back"
        '
        'NextBTN
        '
        Me.NextBTN.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.NextBTN.Location = New System.Drawing.Point(858, 4)
        Me.NextBTN.Name = "NextBTN"
        Me.NextBTN.TabIndex = 2
        Me.NextBTN.Text = "Next >"
        '
        'DoneBTN
        '
        Me.DoneBTN.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DoneBTN.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.DoneBTN.Location = New System.Drawing.Point(937, 4)
        Me.DoneBTN.Name = "DoneBTN"
        Me.DoneBTN.TabIndex = 0
        Me.DoneBTN.Text = "Done"
        '
        'CancelBTN
        '
        Me.CancelBTN.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CancelBTN.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelBTN.Location = New System.Drawing.Point(937, 4)
        Me.CancelBTN.Name = "CancelBTN"
        Me.CancelBTN.TabIndex = 4
        Me.CancelBTN.Text = "Cancel"
        '
        'BrowseCTRL
        '
        Me.BrowseCTRL.Dock = System.Windows.Forms.DockStyle.Fill
        Me.BrowseCTRL.Location = New System.Drawing.Point(4, 4)
        Me.BrowseCTRL.Name = "BrowseCTRL"
        Me.BrowseCTRL.Size = New System.Drawing.Size(296, 326)
        Me.BrowseCTRL.TabIndex = 3
        '
        'LeftPN
        '
        Me.LeftPN.Controls.Add(Me.BrowseCTRL)
        Me.LeftPN.Dock = System.Windows.Forms.DockStyle.Left
        Me.LeftPN.DockPadding.Bottom = 4
        Me.LeftPN.DockPadding.Left = 4
        Me.LeftPN.DockPadding.Top = 4
        Me.LeftPN.Location = New System.Drawing.Point(0, 0)
        Me.LeftPN.Name = "LeftPN"
        Me.LeftPN.Size = New System.Drawing.Size(300, 334)
        Me.LeftPN.TabIndex = 4
        '
        'SplitterV
        '
        Me.SplitterV.Location = New System.Drawing.Point(300, 0)
        Me.SplitterV.Name = "SplitterV"
        Me.SplitterV.Size = New System.Drawing.Size(3, 334)
        Me.SplitterV.TabIndex = 5
        Me.SplitterV.TabStop = False
        '
        'RightPN
        '
        Me.RightPN.Controls.Add(Me.ItemsCTRL)
        Me.RightPN.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RightPN.DockPadding.Bottom = 4
        Me.RightPN.DockPadding.Right = 4
        Me.RightPN.DockPadding.Top = 4
        Me.RightPN.Location = New System.Drawing.Point(303, 0)
        Me.RightPN.Name = "RightPN"
        Me.RightPN.Size = New System.Drawing.Size(713, 334)
        Me.RightPN.TabIndex = 6
        '
        'ItemsCTRL
        '
        Me.ItemsCTRL.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ItemsCTRL.Location = New System.Drawing.Point(0, 4)
        Me.ItemsCTRL.Name = "ItemsCTRL"
        Me.ItemsCTRL.Size = New System.Drawing.Size(709, 326)
        Me.ItemsCTRL.TabIndex = 0
        '
        'AddItemsDlg
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1016, 366)
        Me.Controls.Add(Me.RightPN)
        Me.Controls.Add(Me.SplitterV)
        Me.Controls.Add(Me.LeftPN)
        Me.Controls.Add(Me.BottomPN)
        Me.Name = "AddItemsDlg"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Add Items"
        Me.BottomPN.ResumeLayout(False)
        Me.LeftPN.ResumeLayout(False)
        Me.RightPN.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private m_aServer As Server
    Private m_aGroup As Group
    Private m_pItems As OPCItems

    Public Overloads Function ShowDialog(ByRef pServer As Server, ByRef pGroup As OPCGroup) As Boolean

        ShowDialog = False

        If pServer Is Nothing Or pGroup Is Nothing Then
            Exit Function
        End If

        m_aServer = pServer
        m_aGroup = pServer.GetGroup(pGroup)
        m_pItems = pGroup.OPCItems

        BackBTN.Visible = False
        NextBTN.Visible = True
        CancelBTN.Visible = True
        DoneBTN.Visible = False

        'initialize controls
        BrowseCTRL.Connect(pGroup.Parent())
        ItemsCTRL.Initialize(m_pItems, ItemsCTRL.Mode.Subscribe)

        'show dialog
        If Not ShowDialog() = DialogResult.OK Then
            Exit Function
        End If

        ShowDialog = True

    End Function

    Private Function DoAddItems() As Boolean

        DoAddItems = False

        'get collection of items to add.
        Dim items As Item() = ItemsCTRL.GetItems()

        If items.Length <= 1 Then
            Exit Function
        End If

        'the first item sets the defaults for the collection.
        m_pItems.DefaultIsActive = items(0).Active
        m_pItems.DefaultRequestedDataType = Opc.Type.GetType(items(0).ReqType)
        m_pItems.DefaultAccessPath = items(0).AccessPath

        'assign client handles
        For Each item As Item In items
            item.ClientHandle = m_aServer.NextHandle()
        Next

        'only one item to add.
        If items.Length = 2 Then

            'update the defaults
            m_pItems.DefaultIsActive = items(1).Active

            If Not items(1).ReqType Is Nothing Then
                m_pItems.DefaultRequestedDataType = Opc.Type.GetType(items(1).ReqType)
            End If

            If Not items(1).AccessPath Is Nothing Then
                m_pItems.DefaultAccessPath = items(1).AccessPath
            End If

            'add the item.
            Try

                Dim pItem As OPCItem = m_pItems.AddItem(items(1).ItemID, items(1).ClientHandle)

                Dim results As ArrayList = New ArrayList
                results.Add(Nothing)

                results.Add(item.NewItem(pItem))
                m_aGroup.AddItem(DirectCast(results(1), Item))

                Dim errors(1) As Int32
                ItemsCTRL.Initialize(results.ToArray(GetType(Item)), errors, ItemsCTRL.Mode.Subscribe)
                DoAddItems = True

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Exit Function

        End If

        'initialize arrays for multiple items
        Dim NumItems As Int32 = items.Length - 1

        Dim itemIDs(NumItems) As String
        Dim clientHandles(NumItems) As Int32
        Dim reqTypes(NumItems) As Int16
        Dim accessPaths(NumItems) As String

        For ii As Int32 = 1 To NumItems
            itemIDs(ii) = items(ii).ItemID
            clientHandles(ii) = items(ii).ClientHandle
            reqTypes(ii) = Opc.Type.GetType(items(ii).ReqType)
            accessPaths(ii) = items(ii).AccessPath

            'override the requested active flag since this cannot be set for individual items.
            items(ii).Active = items(0).Active
        Next

        'assign default value for unspecified req types.
        Dim necessary As Boolean = False
        Dim reqType As Int16 = Opc.Type.GetType(items(0).ReqType)

        For ii As Int32 = 1 To NumItems
            If Not items(ii).ReqType Is Nothing Then
                If reqType <> reqTypes(ii) Then
                    necessary = True
                End If
            Else
                reqTypes(ii) = reqType
            End If
        Next

        If Not necessary Then
            reqTypes = Nothing
        End If

        'assign default value for unspecified access paths.
        necessary = False
        Dim accessPath As String = items(0).AccessPath

        For ii As Int32 = 1 To NumItems
            If Not items(ii).AccessPath Is Nothing Then
                If accessPath <> accessPaths(ii) Then
                    necessary = True
                End If
            Else
                accessPaths(ii) = accessPath
            End If
        Next

        If Not necessary Then
            accessPaths = Nothing
        End If

        'call the server.
        Try

            Dim serverHandles As Array
            Dim errors As Array

            m_pItems.AddItems(NumItems, itemIDs, clientHandles, serverHandles, errors, reqTypes, accessPaths)

            Dim results As ArrayList = New ArrayList
            results.Add(Nothing)

            For ii As Long = 1 To serverHandles.Length

                If errors(ii) >= 0 Then
                    results.Add(item.NewItem(m_pItems.GetOPCItem(serverHandles(ii))))
                    m_aGroup.AddItem(DirectCast(results(ii), Item))
                Else
                    results.Add(items(ii))
                End If

            Next

            ItemsCTRL.Initialize(results.ToArray(GetType(Item)), errors, ItemsCTRL.Mode.Subscribe)

            DoAddItems = True

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    Private Function UnDoAddItems() As Boolean

        'get collection of items to remove.
        Dim items As Item() = ItemsCTRL.GetItems()

        'initialize arrays for multiple items
        Dim NumItems As Int32 = items.Length
        Dim serverHandles(NumItems) As Int32

        For ii As Int32 = 1 To NumItems
            serverHandles(ii) = items(ii - 1).ServerHandle
        Next

        Try
            Dim errors As Array
            m_pItems.Remove(NumItems, serverHandles, errors)

            For Each item As Item In items
                m_aGroup.RemoveItem(item.ClientHandle)
            Next

            ItemsCTRL.Initialize(m_pItems, ItemsCTRL.Mode.Subscribe)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    Private Sub ItemPicked(ByVal itemID As String) Handles BrowseCTRL.ItemPicked
        ItemsCTRL.AddItem(itemID)
    End Sub

    Private Sub NextBTN_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NextBTN.Click

        If DoAddItems() Then
            BackBTN.Visible = True
            NextBTN.Visible = False
            CancelBTN.Visible = False
            DoneBTN.Visible = True
        End If

    End Sub

    Private Sub BackBTN_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BackBTN.Click

        UnDoAddItems()

        BackBTN.Visible = False
        NextBTN.Visible = True
        CancelBTN.Visible = True
        DoneBTN.Visible = False

    End Sub
End Class
