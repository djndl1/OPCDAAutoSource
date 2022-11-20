'============================================================================
' TITLE: AddGroupDlg.vb
'
' CONTENTS:
' 
' A dialog used to add a group to a server.
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
Imports System.Globalization

Public Class AddGroupDlg
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
    Friend WithEvents UpdateRateCTRL As System.Windows.Forms.NumericUpDown
    Friend WithEvents UpdateRateLB As System.Windows.Forms.Label
    Friend WithEvents DeadbandCTRL As System.Windows.Forms.NumericUpDown
    Friend WithEvents DeadbandLB As System.Windows.Forms.Label
    Friend WithEvents TimeBiasCTRL As System.Windows.Forms.NumericUpDown
    Friend WithEvents TimeBiasLB As System.Windows.Forms.Label
    Friend WithEvents ActiveCB As System.Windows.Forms.CheckBox
    Friend WithEvents NameTB As System.Windows.Forms.TextBox
    Friend WithEvents NameLB As System.Windows.Forms.Label
    Friend WithEvents ActiveLB As System.Windows.Forms.Label
    Friend WithEvents LocaleTB As System.Windows.Forms.TextBox
    Friend WithEvents LocaleLB As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.BottomPN = New System.Windows.Forms.Panel()
        Me.CancelBTN = New System.Windows.Forms.Button()
        Me.OkBTN = New System.Windows.Forms.Button()
        Me.UpdateRateCTRL = New System.Windows.Forms.NumericUpDown()
        Me.UpdateRateLB = New System.Windows.Forms.Label()
        Me.DeadbandCTRL = New System.Windows.Forms.NumericUpDown()
        Me.DeadbandLB = New System.Windows.Forms.Label()
        Me.TimeBiasCTRL = New System.Windows.Forms.NumericUpDown()
        Me.TimeBiasLB = New System.Windows.Forms.Label()
        Me.ActiveCB = New System.Windows.Forms.CheckBox()
        Me.NameTB = New System.Windows.Forms.TextBox()
        Me.NameLB = New System.Windows.Forms.Label()
        Me.ActiveLB = New System.Windows.Forms.Label()
        Me.LocaleTB = New System.Windows.Forms.TextBox()
        Me.LocaleLB = New System.Windows.Forms.Label()
        Me.BottomPN.SuspendLayout()
        CType(Me.UpdateRateCTRL, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DeadbandCTRL, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TimeBiasCTRL, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BottomPN
        '
        Me.BottomPN.Controls.Add(Me.CancelBTN)
        Me.BottomPN.Controls.Add(Me.OkBTN)
        Me.BottomPN.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.BottomPN.Location = New System.Drawing.Point(0, 148)
        Me.BottomPN.Name = "BottomPN"
        Me.BottomPN.Size = New System.Drawing.Size(276, 34)
        Me.BottomPN.TabIndex = 1
        '
        'CancelBTN
        '
        Me.CancelBTN.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CancelBTN.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelBTN.Location = New System.Drawing.Point(181, 4)
        Me.CancelBTN.Name = "CancelBTN"
        Me.CancelBTN.Size = New System.Drawing.Size(90, 25)
        Me.CancelBTN.TabIndex = 0
        Me.CancelBTN.Text = "Cancel"
        '
        'OkBTN
        '
        Me.OkBTN.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OkBTN.Location = New System.Drawing.Point(5, 4)
        Me.OkBTN.Name = "OkBTN"
        Me.OkBTN.Size = New System.Drawing.Size(90, 25)
        Me.OkBTN.TabIndex = 1
        Me.OkBTN.Text = "OK"
        '
        'UpdateRateCTRL
        '
        Me.UpdateRateCTRL.Location = New System.Drawing.Point(96, 56)
        Me.UpdateRateCTRL.Maximum = New Decimal(New Integer() {-1, 0, 0, 0})
        Me.UpdateRateCTRL.Name = "UpdateRateCTRL"
        Me.UpdateRateCTRL.Size = New System.Drawing.Size(125, 21)
        Me.UpdateRateCTRL.TabIndex = 30
        '
        'UpdateRateLB
        '
        Me.UpdateRateLB.Location = New System.Drawing.Point(5, 56)
        Me.UpdateRateLB.Name = "UpdateRateLB"
        Me.UpdateRateLB.Size = New System.Drawing.Size(91, 25)
        Me.UpdateRateLB.TabIndex = 29
        Me.UpdateRateLB.Text = "Update Rate"
        Me.UpdateRateLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'DeadbandCTRL
        '
        Me.DeadbandCTRL.Location = New System.Drawing.Point(96, 82)
        Me.DeadbandCTRL.Name = "DeadbandCTRL"
        Me.DeadbandCTRL.Size = New System.Drawing.Size(67, 21)
        Me.DeadbandCTRL.TabIndex = 28
        '
        'DeadbandLB
        '
        Me.DeadbandLB.Location = New System.Drawing.Point(5, 82)
        Me.DeadbandLB.Name = "DeadbandLB"
        Me.DeadbandLB.Size = New System.Drawing.Size(91, 25)
        Me.DeadbandLB.TabIndex = 27
        Me.DeadbandLB.Text = "Deadband"
        Me.DeadbandLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TimeBiasCTRL
        '
        Me.TimeBiasCTRL.Location = New System.Drawing.Point(96, 108)
        Me.TimeBiasCTRL.Maximum = New Decimal(New Integer() {16, 0, 0, 0})
        Me.TimeBiasCTRL.Minimum = New Decimal(New Integer() {16, 0, 0, -2147483648})
        Me.TimeBiasCTRL.Name = "TimeBiasCTRL"
        Me.TimeBiasCTRL.Size = New System.Drawing.Size(67, 21)
        Me.TimeBiasCTRL.TabIndex = 26
        Me.TimeBiasCTRL.Value = New Decimal(New Integer() {16, 0, 0, -2147483648})
        '
        'TimeBiasLB
        '
        Me.TimeBiasLB.Location = New System.Drawing.Point(5, 108)
        Me.TimeBiasLB.Name = "TimeBiasLB"
        Me.TimeBiasLB.Size = New System.Drawing.Size(91, 24)
        Me.TimeBiasLB.TabIndex = 25
        Me.TimeBiasLB.Text = "Time Bais"
        Me.TimeBiasLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ActiveCB
        '
        Me.ActiveCB.Location = New System.Drawing.Point(96, 30)
        Me.ActiveCB.Name = "ActiveCB"
        Me.ActiveCB.Size = New System.Drawing.Size(19, 26)
        Me.ActiveCB.TabIndex = 24
        '
        'NameTB
        '
        Me.NameTB.Location = New System.Drawing.Point(96, 4)
        Me.NameTB.Name = "NameTB"
        Me.NameTB.Size = New System.Drawing.Size(230, 21)
        Me.NameTB.TabIndex = 22
        '
        'NameLB
        '
        Me.NameLB.Location = New System.Drawing.Point(5, 4)
        Me.NameLB.Name = "NameLB"
        Me.NameLB.Size = New System.Drawing.Size(91, 25)
        Me.NameLB.TabIndex = 21
        Me.NameLB.Text = "Name"
        Me.NameLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ActiveLB
        '
        Me.ActiveLB.Location = New System.Drawing.Point(5, 30)
        Me.ActiveLB.Name = "ActiveLB"
        Me.ActiveLB.Size = New System.Drawing.Size(91, 25)
        Me.ActiveLB.TabIndex = 23
        Me.ActiveLB.Text = "Active"
        Me.ActiveLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LocaleTB
        '
        Me.LocaleTB.Location = New System.Drawing.Point(96, 134)
        Me.LocaleTB.Name = "LocaleTB"
        Me.LocaleTB.Size = New System.Drawing.Size(67, 21)
        Me.LocaleTB.TabIndex = 32
        '
        'LocaleLB
        '
        Me.LocaleLB.Location = New System.Drawing.Point(5, 134)
        Me.LocaleLB.Name = "LocaleLB"
        Me.LocaleLB.Size = New System.Drawing.Size(91, 24)
        Me.LocaleLB.TabIndex = 31
        Me.LocaleLB.Text = "Locale"
        Me.LocaleLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'AddGroupDlg
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.ClientSize = New System.Drawing.Size(276, 182)
        Me.Controls.Add(Me.LocaleTB)
        Me.Controls.Add(Me.LocaleLB)
        Me.Controls.Add(Me.UpdateRateCTRL)
        Me.Controls.Add(Me.UpdateRateLB)
        Me.Controls.Add(Me.DeadbandCTRL)
        Me.Controls.Add(Me.DeadbandLB)
        Me.Controls.Add(Me.TimeBiasCTRL)
        Me.Controls.Add(Me.TimeBiasLB)
        Me.Controls.Add(Me.ActiveCB)
        Me.Controls.Add(Me.NameTB)
        Me.Controls.Add(Me.NameLB)
        Me.Controls.Add(Me.ActiveLB)
        Me.Controls.Add(Me.BottomPN)
        Me.Name = "AddGroupDlg"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Add Group"
        Me.BottomPN.ResumeLayout(False)
        CType(Me.UpdateRateCTRL, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DeadbandCTRL, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TimeBiasCTRL, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Overloads Function ShowDialog(ByRef pServer As OPCServer) As OPCGroup

        ShowDialog = Nothing

        If pServer Is Nothing Then
            Exit Function
        End If

        Dim pGroups As OPCGroups = pServer.OPCGroups

        NameTB.Text = ""
        ActiveCB.Checked = pGroups.DefaultGroupIsActive
        UpdateRateCTRL.Value = pGroups.DefaultGroupUpdateRate
        DeadbandCTRL.Value = pGroups.DefaultGroupDeadband
        TimeBiasCTRL.Value = pGroups.DefaultGroupTimeBias

        Dim localeID As Int32 = pGroups.DefaultGroupLocaleID

        If localeID <> &H200 And localeID <> &H400 Then
            Try
                LocaleTB.Text = New CultureInfo(localeID).Name
            Catch ex As Exception
                LocaleTB.Text = CultureInfo.CurrentUICulture.Name
            End Try
        Else
            LocaleTB.Text = CultureInfo.CurrentUICulture.Name
        End If

        If Not ShowDialog() = DialogResult.OK Then
            Exit Function
        End If

        pGroups.DefaultGroupIsActive = ActiveCB.Checked
        pGroups.DefaultGroupUpdateRate = UpdateRateCTRL.Value
        pGroups.DefaultGroupDeadband = DeadbandCTRL.Value
        pGroups.DefaultGroupTimeBias = TimeBiasCTRL.Value
        pGroups.DefaultGroupLocaleID = New CultureInfo(LocaleTB.Text).LCID

        Dim pGroup As OPCGroup = pGroups.Add(NameTB.Text)

        If ActiveCB.Checked Then
            pGroup.IsSubscribed = True
        End If

        Dim test As OPCGroup = pGroups.Item(pGroup.Name)

        If test.ServerHandle <> pGroup.ServerHandle Then
            MsgBox("Lookup group by name failed.")
        End If

        test = pGroups.GetOPCGroup(pGroup.ServerHandle)

        If test.ServerHandle <> pGroup.ServerHandle Then
            MsgBox("Lookup group by server handle failed.")
        End If

        test = pGroups.GetOPCGroup(pGroup.Name)

        If test.ServerHandle <> pGroup.ServerHandle Then
            MsgBox("Lookup group by name failed.")
        End If

        ShowDialog = pGroup

    End Function

End Class
