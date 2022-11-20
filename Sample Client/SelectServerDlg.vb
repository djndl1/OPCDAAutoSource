'============================================================================
' TITLE: SelectServerDlg.vb
'
' CONTENTS:
' 
' A dialog used to select a server to connect to.
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

Public Class SelectServerDlg
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
    Friend WithEvents CancelBTN As System.Windows.Forms.Button
    Friend WithEvents OkBTN As System.Windows.Forms.Button
    Friend WithEvents ServersLB As System.Windows.Forms.ListBox
    Friend WithEvents MainPN As System.Windows.Forms.Panel
    Friend WithEvents TopPN As System.Windows.Forms.Panel
    Friend WithEvents NodeLB As System.Windows.Forms.Label
    Friend WithEvents NodeCB As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.BottomPN = New System.Windows.Forms.Panel
        Me.CancelBTN = New System.Windows.Forms.Button
        Me.OkBTN = New System.Windows.Forms.Button
        Me.ServersLB = New System.Windows.Forms.ListBox
        Me.MainPN = New System.Windows.Forms.Panel
        Me.TopPN = New System.Windows.Forms.Panel
        Me.NodeCB = New System.Windows.Forms.ComboBox
        Me.NodeLB = New System.Windows.Forms.Label
        Me.BottomPN.SuspendLayout()
        Me.MainPN.SuspendLayout()
        Me.TopPN.SuspendLayout()
        Me.SuspendLayout()
        '
        'BottomPN
        '
        Me.BottomPN.Controls.Add(Me.CancelBTN)
        Me.BottomPN.Controls.Add(Me.OkBTN)
        Me.BottomPN.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.BottomPN.Location = New System.Drawing.Point(0, 246)
        Me.BottomPN.Name = "BottomPN"
        Me.BottomPN.Size = New System.Drawing.Size(256, 32)
        Me.BottomPN.TabIndex = 1
        '
        'CancelBTN
        '
        Me.CancelBTN.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CancelBTN.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelBTN.Location = New System.Drawing.Point(177, 4)
        Me.CancelBTN.Name = "CancelBTN"
        Me.CancelBTN.TabIndex = 0
        Me.CancelBTN.Text = "Cancel"
        '
        'OkBTN
        '
        Me.OkBTN.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OkBTN.Location = New System.Drawing.Point(4, 4)
        Me.OkBTN.Name = "OkBTN"
        Me.OkBTN.TabIndex = 1
        Me.OkBTN.Text = "OK"
        '
        'ServersLB
        '
        Me.ServersLB.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ServersLB.Location = New System.Drawing.Point(4, 4)
        Me.ServersLB.Name = "ServersLB"
        Me.ServersLB.Size = New System.Drawing.Size(248, 199)
        Me.ServersLB.TabIndex = 2
        '
        'MainPN
        '
        Me.MainPN.Controls.Add(Me.ServersLB)
        Me.MainPN.Dock = System.Windows.Forms.DockStyle.Fill
        Me.MainPN.DockPadding.All = 4
        Me.MainPN.Location = New System.Drawing.Point(0, 32)
        Me.MainPN.Name = "MainPN"
        Me.MainPN.Size = New System.Drawing.Size(256, 214)
        Me.MainPN.TabIndex = 3
        '
        'TopPN
        '
        Me.TopPN.Controls.Add(Me.NodeCB)
        Me.TopPN.Controls.Add(Me.NodeLB)
        Me.TopPN.Dock = System.Windows.Forms.DockStyle.Top
        Me.TopPN.Location = New System.Drawing.Point(0, 0)
        Me.TopPN.Name = "TopPN"
        Me.TopPN.Size = New System.Drawing.Size(256, 32)
        Me.TopPN.TabIndex = 4
        '
        'NodeCB
        '
        Me.NodeCB.Location = New System.Drawing.Point(48, 5)
        Me.NodeCB.Name = "NodeCB"
        Me.NodeCB.Size = New System.Drawing.Size(204, 21)
        Me.NodeCB.TabIndex = 4
        '
        'NodeLB
        '
        Me.NodeLB.Location = New System.Drawing.Point(4, 4)
        Me.NodeLB.Name = "NodeLB"
        Me.NodeLB.Size = New System.Drawing.Size(40, 23)
        Me.NodeLB.TabIndex = 3
        Me.NodeLB.Text = "Node"
        Me.NodeLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'SelectServerDlg
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(256, 278)
        Me.Controls.Add(Me.MainPN)
        Me.Controls.Add(Me.TopPN)
        Me.Controls.Add(Me.BottomPN)
        Me.Name = "SelectServerDlg"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Select Server"
        Me.BottomPN.ResumeLayout(False)
        Me.MainPN.ResumeLayout(False)
        Me.TopPN.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private m_pServer As OPCServer

    Public Overloads Function ShowDialog(ByRef pServer As OPCServer, ByRef node As String) As String

        ShowDialog = Nothing

        m_pServer = pServer

        NodeCB.Items.Clear()

        NodeCB.Items.Add("")
        NodeCB.SelectedItem = ""

        NodeCB.Items.AddRange(Opc.Interop.EnumComputers())
        NodeCB.SelectedItem = node

        'show dialog
        If Not ShowDialog() = DialogResult.OK Then
            Exit Function
        End If

        node = NodeCB.SelectedItem
        ShowDialog = ServersLB.SelectedItem

    End Function

    Private Sub ServersLB_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ServersLB.DoubleClick
        DialogResult = DialogResult.OK
    End Sub

    Private Sub NodeCB_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles NodeCB.SelectedIndexChanged

        Try

            ServersLB.Items.Clear()

            Dim servers As Array = m_pServer.GetOPCServers(NodeCB.SelectedItem)

            For Each server As String In servers
                ServersLB.Items.Add(server)
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
End Class
