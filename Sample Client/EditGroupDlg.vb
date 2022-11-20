'============================================================================
' TITLE: EditGroupDlg.vb
'
' CONTENTS:
' 
' A dialog used to edit the state of a group.
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

Public Class EditGroupDlg
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
    Friend WithEvents NameTB As System.Windows.Forms.TextBox
    Friend WithEvents ActiveLB As System.Windows.Forms.Label
    Friend WithEvents ActiveCB As System.Windows.Forms.CheckBox
    Friend WithEvents SubscribedCB As System.Windows.Forms.CheckBox
    Friend WithEvents SubscribedLB As System.Windows.Forms.Label
    Friend WithEvents ClienHandleLB As System.Windows.Forms.Label
    Friend WithEvents ServerHandleTB As System.Windows.Forms.TextBox
    Friend WithEvents ServerHandleLB As System.Windows.Forms.Label
    Friend WithEvents LocaleTB As System.Windows.Forms.TextBox
    Friend WithEvents LocaleLB As System.Windows.Forms.Label
    Friend WithEvents TimeBiasLB As System.Windows.Forms.Label
    Friend WithEvents TimeBiasCTRL As System.Windows.Forms.NumericUpDown
    Friend WithEvents DeadbandCTRL As System.Windows.Forms.NumericUpDown
    Friend WithEvents DeadbandLB As System.Windows.Forms.Label
    Friend WithEvents UpdateRateCTRL As System.Windows.Forms.NumericUpDown
    Friend WithEvents UpdateRateLB As System.Windows.Forms.Label
    Friend WithEvents NameLB As System.Windows.Forms.Label
    Friend WithEvents ClientHandleTB As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.BottomPN = New System.Windows.Forms.Panel
        Me.CancelBTN = New System.Windows.Forms.Button
        Me.OkBTN = New System.Windows.Forms.Button
        Me.NameTB = New System.Windows.Forms.TextBox
        Me.NameLB = New System.Windows.Forms.Label
        Me.ActiveLB = New System.Windows.Forms.Label
        Me.ActiveCB = New System.Windows.Forms.CheckBox
        Me.SubscribedCB = New System.Windows.Forms.CheckBox
        Me.SubscribedLB = New System.Windows.Forms.Label
        Me.ClientHandleTB = New System.Windows.Forms.TextBox
        Me.ClienHandleLB = New System.Windows.Forms.Label
        Me.ServerHandleTB = New System.Windows.Forms.TextBox
        Me.ServerHandleLB = New System.Windows.Forms.Label
        Me.LocaleTB = New System.Windows.Forms.TextBox
        Me.LocaleLB = New System.Windows.Forms.Label
        Me.TimeBiasLB = New System.Windows.Forms.Label
        Me.TimeBiasCTRL = New System.Windows.Forms.NumericUpDown
        Me.DeadbandCTRL = New System.Windows.Forms.NumericUpDown
        Me.DeadbandLB = New System.Windows.Forms.Label
        Me.UpdateRateCTRL = New System.Windows.Forms.NumericUpDown
        Me.UpdateRateLB = New System.Windows.Forms.Label
        Me.BottomPN.SuspendLayout()
        CType(Me.TimeBiasCTRL, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DeadbandCTRL, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UpdateRateCTRL, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BottomPN
        '
        Me.BottomPN.Controls.Add(Me.CancelBTN)
        Me.BottomPN.Controls.Add(Me.OkBTN)
        Me.BottomPN.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.BottomPN.Location = New System.Drawing.Point(0, 218)
        Me.BottomPN.Name = "BottomPN"
        Me.BottomPN.Size = New System.Drawing.Size(276, 32)
        Me.BottomPN.TabIndex = 0
        '
        'CancelBTN
        '
        Me.CancelBTN.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CancelBTN.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelBTN.Location = New System.Drawing.Point(197, 4)
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
        'NameTB
        '
        Me.NameTB.Location = New System.Drawing.Point(80, 4)
        Me.NameTB.Name = "NameTB"
        Me.NameTB.Size = New System.Drawing.Size(192, 20)
        Me.NameTB.TabIndex = 2
        Me.NameTB.Text = ""
        '
        'NameLB
        '
        Me.NameLB.Location = New System.Drawing.Point(4, 4)
        Me.NameLB.Name = "NameLB"
        Me.NameLB.Size = New System.Drawing.Size(76, 23)
        Me.NameLB.TabIndex = 1
        Me.NameLB.Text = "Name"
        Me.NameLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ActiveLB
        '
        Me.ActiveLB.Location = New System.Drawing.Point(4, 28)
        Me.ActiveLB.Name = "ActiveLB"
        Me.ActiveLB.Size = New System.Drawing.Size(76, 23)
        Me.ActiveLB.TabIndex = 3
        Me.ActiveLB.Text = "Active"
        Me.ActiveLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ActiveCB
        '
        Me.ActiveCB.Location = New System.Drawing.Point(80, 28)
        Me.ActiveCB.Name = "ActiveCB"
        Me.ActiveCB.Size = New System.Drawing.Size(16, 24)
        Me.ActiveCB.TabIndex = 4
        '
        'SubscribedCB
        '
        Me.SubscribedCB.Location = New System.Drawing.Point(80, 52)
        Me.SubscribedCB.Name = "SubscribedCB"
        Me.SubscribedCB.Size = New System.Drawing.Size(16, 24)
        Me.SubscribedCB.TabIndex = 8
        '
        'SubscribedLB
        '
        Me.SubscribedLB.Location = New System.Drawing.Point(4, 52)
        Me.SubscribedLB.Name = "SubscribedLB"
        Me.SubscribedLB.Size = New System.Drawing.Size(76, 23)
        Me.SubscribedLB.TabIndex = 7
        Me.SubscribedLB.Text = "Subscribed"
        Me.SubscribedLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ClientHandleTB
        '
        Me.ClientHandleTB.Location = New System.Drawing.Point(80, 196)
        Me.ClientHandleTB.Name = "ClientHandleTB"
        Me.ClientHandleTB.ReadOnly = True
        Me.ClientHandleTB.Size = New System.Drawing.Size(104, 20)
        Me.ClientHandleTB.TabIndex = 10
        Me.ClientHandleTB.Text = ""
        '
        'ClienHandleLB
        '
        Me.ClienHandleLB.Location = New System.Drawing.Point(4, 196)
        Me.ClienHandleLB.Name = "ClienHandleLB"
        Me.ClienHandleLB.Size = New System.Drawing.Size(76, 23)
        Me.ClienHandleLB.TabIndex = 9
        Me.ClienHandleLB.Text = "Client Handle"
        Me.ClienHandleLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ServerHandleTB
        '
        Me.ServerHandleTB.Location = New System.Drawing.Point(80, 172)
        Me.ServerHandleTB.Name = "ServerHandleTB"
        Me.ServerHandleTB.ReadOnly = True
        Me.ServerHandleTB.Size = New System.Drawing.Size(104, 20)
        Me.ServerHandleTB.TabIndex = 12
        Me.ServerHandleTB.Text = ""
        '
        'ServerHandleLB
        '
        Me.ServerHandleLB.Location = New System.Drawing.Point(4, 172)
        Me.ServerHandleLB.Name = "ServerHandleLB"
        Me.ServerHandleLB.Size = New System.Drawing.Size(76, 23)
        Me.ServerHandleLB.TabIndex = 11
        Me.ServerHandleLB.Text = "Server Handle"
        Me.ServerHandleLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LocaleTB
        '
        Me.LocaleTB.Location = New System.Drawing.Point(80, 148)
        Me.LocaleTB.Name = "LocaleTB"
        Me.LocaleTB.Size = New System.Drawing.Size(56, 20)
        Me.LocaleTB.TabIndex = 14
        Me.LocaleTB.Text = ""
        '
        'LocaleLB
        '
        Me.LocaleLB.Location = New System.Drawing.Point(4, 148)
        Me.LocaleLB.Name = "LocaleLB"
        Me.LocaleLB.Size = New System.Drawing.Size(76, 23)
        Me.LocaleLB.TabIndex = 13
        Me.LocaleLB.Text = "Locale"
        Me.LocaleLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TimeBiasLB
        '
        Me.TimeBiasLB.Location = New System.Drawing.Point(4, 124)
        Me.TimeBiasLB.Name = "TimeBiasLB"
        Me.TimeBiasLB.Size = New System.Drawing.Size(76, 23)
        Me.TimeBiasLB.TabIndex = 15
        Me.TimeBiasLB.Text = "Time Bais"
        Me.TimeBiasLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TimeBiasCTRL
        '
        Me.TimeBiasCTRL.Location = New System.Drawing.Point(80, 124)
        Me.TimeBiasCTRL.Maximum = New Decimal(New Integer() {16, 0, 0, 0})
        Me.TimeBiasCTRL.Minimum = New Decimal(New Integer() {16, 0, 0, -2147483648})
        Me.TimeBiasCTRL.Name = "TimeBiasCTRL"
        Me.TimeBiasCTRL.Size = New System.Drawing.Size(56, 20)
        Me.TimeBiasCTRL.TabIndex = 16
        Me.TimeBiasCTRL.Value = New Decimal(New Integer() {16, 0, 0, -2147483648})
        '
        'DeadbandCTRL
        '
        Me.DeadbandCTRL.Location = New System.Drawing.Point(80, 100)
        Me.DeadbandCTRL.Name = "DeadbandCTRL"
        Me.DeadbandCTRL.Size = New System.Drawing.Size(56, 20)
        Me.DeadbandCTRL.TabIndex = 18
        '
        'DeadbandLB
        '
        Me.DeadbandLB.Location = New System.Drawing.Point(4, 100)
        Me.DeadbandLB.Name = "DeadbandLB"
        Me.DeadbandLB.Size = New System.Drawing.Size(76, 23)
        Me.DeadbandLB.TabIndex = 17
        Me.DeadbandLB.Text = "Deadband"
        Me.DeadbandLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'UpdateRateCTRL
        '
        Me.UpdateRateCTRL.Location = New System.Drawing.Point(80, 76)
        Me.UpdateRateCTRL.Maximum = New Decimal(New Integer() {-1, 0, 0, 0})
        Me.UpdateRateCTRL.Name = "UpdateRateCTRL"
        Me.UpdateRateCTRL.Size = New System.Drawing.Size(104, 20)
        Me.UpdateRateCTRL.TabIndex = 20
        '
        'UpdateRateLB
        '
        Me.UpdateRateLB.Location = New System.Drawing.Point(4, 76)
        Me.UpdateRateLB.Name = "UpdateRateLB"
        Me.UpdateRateLB.Size = New System.Drawing.Size(76, 23)
        Me.UpdateRateLB.TabIndex = 19
        Me.UpdateRateLB.Text = "Update Rate"
        Me.UpdateRateLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'EditGroupDlg
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.CancelBTN
        Me.ClientSize = New System.Drawing.Size(276, 250)
        Me.Controls.Add(Me.UpdateRateCTRL)
        Me.Controls.Add(Me.UpdateRateLB)
        Me.Controls.Add(Me.DeadbandCTRL)
        Me.Controls.Add(Me.DeadbandLB)
        Me.Controls.Add(Me.TimeBiasCTRL)
        Me.Controls.Add(Me.TimeBiasLB)
        Me.Controls.Add(Me.LocaleTB)
        Me.Controls.Add(Me.ServerHandleTB)
        Me.Controls.Add(Me.ClientHandleTB)
        Me.Controls.Add(Me.NameTB)
        Me.Controls.Add(Me.LocaleLB)
        Me.Controls.Add(Me.ServerHandleLB)
        Me.Controls.Add(Me.ClienHandleLB)
        Me.Controls.Add(Me.SubscribedCB)
        Me.Controls.Add(Me.SubscribedLB)
        Me.Controls.Add(Me.ActiveCB)
        Me.Controls.Add(Me.NameLB)
        Me.Controls.Add(Me.ActiveLB)
        Me.Controls.Add(Me.BottomPN)
        Me.Name = "EditGroupDlg"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Edit Group"
        Me.BottomPN.ResumeLayout(False)
        CType(Me.TimeBiasCTRL, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DeadbandCTRL, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UpdateRateCTRL, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Overloads Function ShowDialog(ByRef pGroup As OPCGroup) As Boolean

        ShowDialog = False

        If pGroup Is Nothing Then
            Exit Function
        End If

        NameTB.Text = pGroup.Name
        ActiveCB.Checked = pGroup.IsActive
        SubscribedCB.Checked = pGroup.IsSubscribed
        UpdateRateCTRL.Value = pGroup.UpdateRate
        DeadbandCTRL.Value = pGroup.DeadBand
        TimeBiasCTRL.Value = pGroup.TimeBias

        Try
            LocaleTB.Text = New CultureInfo(pGroup.LocaleID).Name
        Catch ex As Exception
            LocaleTB.Text = CultureInfo.CurrentUICulture.Name
        End Try

        ServerHandleTB.Text = String.Format("&H{0:X8}", pGroup.ServerHandle)
        ClientHandleTB.Text = String.Format("&H{0:X8}", pGroup.ClientHandle)

        If Not ShowDialog() = DialogResult.OK Then
            Exit Function
        End If

        pGroup.Name = NameTB.Text
        pGroup.IsActive = ActiveCB.Checked
        pGroup.IsSubscribed = SubscribedCB.Checked
        pGroup.UpdateRate = UpdateRateCTRL.Value
        pGroup.DeadBand = DeadbandCTRL.Value
        pGroup.TimeBias = TimeBiasCTRL.Value

        Try
            pGroup.LocaleID = New CultureInfo(LocaleTB.Text).LCID
        Catch ex As Exception
            'Do nothing
        End Try

        ShowDialog = True

    End Function

End Class
