'============================================================================
' TITLE: WriteItemsDlg.vb
'
' CONTENTS:
' 
' A dialog used to write a set of items to a server.
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

Public Class WriteItemsDlg
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
    Friend WithEvents BackBTN As System.Windows.Forms.Button
    Friend WithEvents NextBTN As System.Windows.Forms.Button
    Friend WithEvents DoneBTN As System.Windows.Forms.Button
    Friend WithEvents CancelBTN As System.Windows.Forms.Button
    Friend WithEvents LeftPN As System.Windows.Forms.Panel
    Friend WithEvents ItemsCTRL As VbNetSampleClient.GroupItemsCtrl
    Friend WithEvents SplitterV As System.Windows.Forms.Splitter
    Friend WithEvents RightPN As System.Windows.Forms.Panel
    Friend WithEvents WriteItemsCTRL As VbNetSampleClient.ItemsCtrl
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.BottomPN = New System.Windows.Forms.Panel
        Me.BackBTN = New System.Windows.Forms.Button
        Me.NextBTN = New System.Windows.Forms.Button
        Me.DoneBTN = New System.Windows.Forms.Button
        Me.CancelBTN = New System.Windows.Forms.Button
        Me.LeftPN = New System.Windows.Forms.Panel
        Me.ItemsCTRL = New VbNetSampleClient.GroupItemsCtrl
        Me.SplitterV = New System.Windows.Forms.Splitter
        Me.RightPN = New System.Windows.Forms.Panel
        Me.WriteItemsCTRL = New VbNetSampleClient.ItemsCtrl
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
        Me.BottomPN.TabIndex = 4
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
        'LeftPN
        '
        Me.LeftPN.Controls.Add(Me.ItemsCTRL)
        Me.LeftPN.Dock = System.Windows.Forms.DockStyle.Left
        Me.LeftPN.DockPadding.Bottom = 4
        Me.LeftPN.DockPadding.Left = 4
        Me.LeftPN.DockPadding.Top = 4
        Me.LeftPN.Location = New System.Drawing.Point(0, 0)
        Me.LeftPN.Name = "LeftPN"
        Me.LeftPN.Size = New System.Drawing.Size(400, 334)
        Me.LeftPN.TabIndex = 6
        '
        'ItemsCTRL
        '
        Me.ItemsCTRL.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ItemsCTRL.Location = New System.Drawing.Point(4, 4)
        Me.ItemsCTRL.Name = "ItemsCTRL"
        Me.ItemsCTRL.Size = New System.Drawing.Size(396, 326)
        Me.ItemsCTRL.TabIndex = 0
        '
        'SplitterV
        '
        Me.SplitterV.Location = New System.Drawing.Point(400, 0)
        Me.SplitterV.Name = "SplitterV"
        Me.SplitterV.Size = New System.Drawing.Size(3, 334)
        Me.SplitterV.TabIndex = 7
        Me.SplitterV.TabStop = False
        '
        'RightPN
        '
        Me.RightPN.Controls.Add(Me.WriteItemsCTRL)
        Me.RightPN.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RightPN.DockPadding.Bottom = 4
        Me.RightPN.DockPadding.Right = 4
        Me.RightPN.DockPadding.Top = 4
        Me.RightPN.Location = New System.Drawing.Point(403, 0)
        Me.RightPN.Name = "RightPN"
        Me.RightPN.Size = New System.Drawing.Size(613, 334)
        Me.RightPN.TabIndex = 8
        '
        'WriteItemsCTRL
        '
        Me.WriteItemsCTRL.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WriteItemsCTRL.Location = New System.Drawing.Point(0, 4)
        Me.WriteItemsCTRL.Name = "WriteItemsCTRL"
        Me.WriteItemsCTRL.Size = New System.Drawing.Size(609, 326)
        Me.WriteItemsCTRL.TabIndex = 0
        '
        'WriteItemsDlg
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1016, 366)
        Me.Controls.Add(Me.RightPN)
        Me.Controls.Add(Me.SplitterV)
        Me.Controls.Add(Me.LeftPN)
        Me.Controls.Add(Me.BottomPN)
        Me.Name = "WriteItemsDlg"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Write Items"
        Me.BottomPN.ResumeLayout(False)
        Me.LeftPN.ResumeLayout(False)
        Me.RightPN.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private m_aServer As Server
    Private m_pGroup As OPCGroup
    Private m_isAsync As Boolean

    Public Overloads Function ShowDialog(ByRef pServer As Server, ByRef pGroup As OPCGroup, ByVal isAsync As Boolean) As Boolean

        ShowDialog = False

        If pServer Is Nothing Or pGroup Is Nothing Then
            Exit Function
        End If

        m_aServer = pServer
        m_pGroup = pGroup
        m_isAsync = isAsync

        BackBTN.Visible = False
        NextBTN.Visible = True
        CancelBTN.Visible = True
        DoneBTN.Visible = False

        'initialize controls
        ItemsCTRL.Initialize(m_pGroup.OPCItems)
        WriteItemsCTRL.Initialize(m_pGroup.OPCItems, WriteItemsCTRL.Mode.Write)

        'show dialog
        If Not ShowDialog() = DialogResult.OK Then
            Exit Function
        End If

        ShowDialog = True

    End Function

    Private Function DoSingleWrite(ByRef items() As Item) As Boolean

        Try

            Dim pItem As OPCItem = m_pGroup.OPCItems.GetOPCItem(items(0).ServerHandle)

            pItem.Write(items(0).Value)

            Dim results As ArrayList = New ArrayList

            results.Add(Nothing)
            results.Add(items(0))

            Dim errors(1) As Int32
            WriteItemsCTRL.Initialize(results.ToArray(GetType(Item)), errors, WriteItemsCTRL.Mode.Write)
            DoSingleWrite = True

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function
    Private Function DoWriteItems() As Boolean

        DoWriteItems = False

        'get collection of items to read.
        Dim items As Item() = WriteItemsCTRL.GetItems()

        For Each item As Item In items

            If Not item.Value Is Nothing Then

                'convert a decimal array to an object array since decimal arrays can't be marshalled by DCOM.
                If item.Value.GetType() Is GetType(Decimal()) Then
                    Dim objects As ArrayList = New ArrayList
                    objects.AddRange(item.Value)
                    item.Value = DirectCast(objects.ToArray(GetType(Object)), Object())
                End If

            End If
        Next

        If items.Length < 1 Then
            Exit Function
        End If

        'only one item to read.
        If items.Length = 1 And Not m_isAsync Then
            DoWriteItems = DoSingleWrite(items)
            Exit Function
        End If

        'initialize arrays for multiple items
        Dim NumItems As Int32 = items.Length

        Dim itemIDs(NumItems) As String
        Dim serverHandles(NumItems) As Int32
        Dim values(NumItems) As Object

        For ii As Int32 = 0 To NumItems - 1
            serverHandles(ii + 1) = items(ii).ServerHandle
            values(ii + 1) = items(ii).Value
        Next

        If Not m_isAsync Then

            Dim errors As Array

            Try
                m_pGroup.SyncWrite(NumItems, serverHandles, values, errors)
            Catch ex As Exception
                MsgBox(ex.Message)
                Exit Function
            End Try


            Dim results As ArrayList = New ArrayList
            results.Add(Nothing)

            For ii As Long = 1 To NumItems
                results.Add(items(ii - 1))
            Next

            WriteItemsCTRL.Initialize(results.ToArray(GetType(Item)), errors, WriteItemsCTRL.Mode.Write)

        Else

            Dim errors As Array
            Dim cancelID As Int32

            Try
                m_pGroup.AsyncWrite(NumItems, serverHandles, values, errors, m_aServer.NextHandle(), cancelID)
            Catch ex As Exception
                MsgBox(ex.Message)
                Exit Function
            End Try

            Dim results As ArrayList = New ArrayList
            results.Add(Nothing)

            For ii As Long = 1 To NumItems
                results.Add(items(ii - 1))
            Next

            WriteItemsCTRL.Initialize(results.ToArray(GetType(Item)), errors, WriteItemsCTRL.Mode.Write)

        End If

        DoWriteItems = True

    End Function

    Private Sub ItemsCTRL_ItemPicked(ByRef item As OPCAutomation.OPCItem) Handles ItemsCTRL.ItemPicked
        WriteItemsCTRL.AddItem(item)
    End Sub

    Private Sub NextBTN_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NextBTN.Click

        If DoWriteItems() Then
            BackBTN.Visible = True
            NextBTN.Visible = False
            CancelBTN.Visible = False
            DoneBTN.Visible = True
        End If

    End Sub

    Private Sub BackBTN_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BackBTN.Click

        WriteItemsCTRL.Initialize(m_pGroup.OPCItems, WriteItemsCTRL.Mode.Write)

        BackBTN.Visible = False
        NextBTN.Visible = True
        CancelBTN.Visible = True
        DoneBTN.Visible = False

    End Sub

End Class
