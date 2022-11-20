'============================================================================
' TITLE: EditItemDlg.vb
'
' CONTENTS:
' 
' A dialog used to edit the state of an item.
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

Public Class EditItemDlg
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
    Friend WithEvents WritePN As System.Windows.Forms.Panel
    Friend WithEvents ValueLB As System.Windows.Forms.Label
    Friend WithEvents SubscribePN As System.Windows.Forms.Panel
    Friend WithEvents ActiveLB As System.Windows.Forms.Label
    Friend WithEvents DataTypeLB As System.Windows.Forms.Label
    Friend WithEvents AccessPathLB As System.Windows.Forms.Label
    Friend WithEvents AccessPathTB As System.Windows.Forms.TextBox
    Friend WithEvents DataTypeCTRL As Opc.Da.SampleClient.DataTypeCtrl
    Friend WithEvents ActiveCB As System.Windows.Forms.CheckBox
    Friend WithEvents ValueCTRL As Opc.Da.SampleClient.ValueCtrl
    Friend WithEvents ItemIDTB As System.Windows.Forms.TextBox
    Friend WithEvents ItemIDLB As System.Windows.Forms.Label
    Friend WithEvents TopPN As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.BottomPN = New System.Windows.Forms.Panel
        Me.CancelBTN = New System.Windows.Forms.Button
        Me.OkBTN = New System.Windows.Forms.Button
        Me.WritePN = New System.Windows.Forms.Panel
        Me.ValueCTRL = New Opc.Da.SampleClient.ValueCtrl
        Me.ValueLB = New System.Windows.Forms.Label
        Me.SubscribePN = New System.Windows.Forms.Panel
        Me.ActiveCB = New System.Windows.Forms.CheckBox
        Me.DataTypeCTRL = New Opc.Da.SampleClient.DataTypeCtrl
        Me.AccessPathTB = New System.Windows.Forms.TextBox
        Me.AccessPathLB = New System.Windows.Forms.Label
        Me.DataTypeLB = New System.Windows.Forms.Label
        Me.ActiveLB = New System.Windows.Forms.Label
        Me.ItemIDTB = New System.Windows.Forms.TextBox
        Me.ItemIDLB = New System.Windows.Forms.Label
        Me.TopPN = New System.Windows.Forms.Panel
        Me.BottomPN.SuspendLayout()
        Me.WritePN.SuspendLayout()
        Me.SubscribePN.SuspendLayout()
        Me.TopPN.SuspendLayout()
        Me.SuspendLayout()
        '
        'BottomPN
        '
        Me.BottomPN.Controls.Add(Me.CancelBTN)
        Me.BottomPN.Controls.Add(Me.OkBTN)
        Me.BottomPN.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.BottomPN.Location = New System.Drawing.Point(0, 102)
        Me.BottomPN.Name = "BottomPN"
        Me.BottomPN.Size = New System.Drawing.Size(324, 32)
        Me.BottomPN.TabIndex = 1
        '
        'CancelBTN
        '
        Me.CancelBTN.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CancelBTN.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelBTN.Location = New System.Drawing.Point(245, 4)
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
        'WritePN
        '
        Me.WritePN.Controls.Add(Me.ValueCTRL)
        Me.WritePN.Controls.Add(Me.ValueLB)
        Me.WritePN.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WritePN.Location = New System.Drawing.Point(0, 24)
        Me.WritePN.Name = "WritePN"
        Me.WritePN.Size = New System.Drawing.Size(324, 78)
        Me.WritePN.TabIndex = 5
        '
        'ValueCTRL
        '
        Me.ValueCTRL.DataValue = Nothing
        Me.ValueCTRL.ItemID = Nothing
        Me.ValueCTRL.Location = New System.Drawing.Point(80, 4)
        Me.ValueCTRL.Name = "ValueCTRL"
        Me.ValueCTRL.Size = New System.Drawing.Size(240, 20)
        Me.ValueCTRL.TabIndex = 5
        '
        'ValueLB
        '
        Me.ValueLB.Location = New System.Drawing.Point(4, 4)
        Me.ValueLB.Name = "ValueLB"
        Me.ValueLB.Size = New System.Drawing.Size(72, 23)
        Me.ValueLB.TabIndex = 4
        Me.ValueLB.Text = "Value"
        Me.ValueLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'SubscribePN
        '
        Me.SubscribePN.Controls.Add(Me.ActiveCB)
        Me.SubscribePN.Controls.Add(Me.DataTypeCTRL)
        Me.SubscribePN.Controls.Add(Me.AccessPathTB)
        Me.SubscribePN.Controls.Add(Me.AccessPathLB)
        Me.SubscribePN.Controls.Add(Me.DataTypeLB)
        Me.SubscribePN.Controls.Add(Me.ActiveLB)
        Me.SubscribePN.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SubscribePN.Location = New System.Drawing.Point(0, 24)
        Me.SubscribePN.Name = "SubscribePN"
        Me.SubscribePN.Size = New System.Drawing.Size(324, 78)
        Me.SubscribePN.TabIndex = 6
        '
        'ActiveCB
        '
        Me.ActiveCB.Location = New System.Drawing.Point(80, 4)
        Me.ActiveCB.Name = "ActiveCB"
        Me.ActiveCB.Size = New System.Drawing.Size(16, 24)
        Me.ActiveCB.TabIndex = 9
        '
        'DataTypeCTRL
        '
        Me.DataTypeCTRL.Location = New System.Drawing.Point(80, 28)
        Me.DataTypeCTRL.Name = "DataTypeCTRL"
        Me.DataTypeCTRL.SelectedType = Nothing
        Me.DataTypeCTRL.Size = New System.Drawing.Size(112, 24)
        Me.DataTypeCTRL.TabIndex = 8
        '
        'AccessPathTB
        '
        Me.AccessPathTB.Location = New System.Drawing.Point(80, 52)
        Me.AccessPathTB.Name = "AccessPathTB"
        Me.AccessPathTB.Size = New System.Drawing.Size(192, 20)
        Me.AccessPathTB.TabIndex = 7
        Me.AccessPathTB.Text = ""
        '
        'AccessPathLB
        '
        Me.AccessPathLB.Location = New System.Drawing.Point(4, 52)
        Me.AccessPathLB.Name = "AccessPathLB"
        Me.AccessPathLB.Size = New System.Drawing.Size(76, 23)
        Me.AccessPathLB.TabIndex = 6
        Me.AccessPathLB.Text = "Access Path"
        Me.AccessPathLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'DataTypeLB
        '
        Me.DataTypeLB.Location = New System.Drawing.Point(4, 28)
        Me.DataTypeLB.Name = "DataTypeLB"
        Me.DataTypeLB.Size = New System.Drawing.Size(76, 23)
        Me.DataTypeLB.TabIndex = 5
        Me.DataTypeLB.Text = "Data Type"
        Me.DataTypeLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ActiveLB
        '
        Me.ActiveLB.Location = New System.Drawing.Point(4, 4)
        Me.ActiveLB.Name = "ActiveLB"
        Me.ActiveLB.Size = New System.Drawing.Size(76, 23)
        Me.ActiveLB.TabIndex = 4
        Me.ActiveLB.Text = "Active"
        Me.ActiveLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ItemIDTB
        '
        Me.ItemIDTB.Location = New System.Drawing.Point(80, 4)
        Me.ItemIDTB.Name = "ItemIDTB"
        Me.ItemIDTB.Size = New System.Drawing.Size(192, 20)
        Me.ItemIDTB.TabIndex = 4
        Me.ItemIDTB.Text = ""
        '
        'ItemIDLB
        '
        Me.ItemIDLB.Location = New System.Drawing.Point(4, 4)
        Me.ItemIDLB.Name = "ItemIDLB"
        Me.ItemIDLB.Size = New System.Drawing.Size(76, 23)
        Me.ItemIDLB.TabIndex = 3
        Me.ItemIDLB.Text = "Item ID"
        Me.ItemIDLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TopPN
        '
        Me.TopPN.Controls.Add(Me.ItemIDTB)
        Me.TopPN.Controls.Add(Me.ItemIDLB)
        Me.TopPN.Dock = System.Windows.Forms.DockStyle.Top
        Me.TopPN.Location = New System.Drawing.Point(0, 0)
        Me.TopPN.Name = "TopPN"
        Me.TopPN.Size = New System.Drawing.Size(324, 24)
        Me.TopPN.TabIndex = 7
        '
        'EditItemDlg
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(324, 134)
        Me.Controls.Add(Me.WritePN)
        Me.Controls.Add(Me.SubscribePN)
        Me.Controls.Add(Me.TopPN)
        Me.Controls.Add(Me.BottomPN)
        Me.Name = "EditItemDlg"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Edit Item"
        Me.BottomPN.ResumeLayout(False)
        Me.WritePN.ResumeLayout(False)
        Me.SubscribePN.ResumeLayout(False)
        Me.TopPN.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Overloads Function ShowDialog(ByRef item As Item, ByVal mode As ItemsCtrl.Mode) As Boolean

        ShowDialog = False

        ItemIDTB.Text = item.ItemID
        ActiveCB.Checked = item.Active
        DataTypeCTRL.SelectedType = item.ReqType
        AccessPathTB.Text = item.AccessPath
        ValueCTRL.DataValue = item.Value

        If mode = ItemsCtrl.Mode.Subscribe Then
            ItemIDTB.ReadOnly = False
            SubscribePN.Visible = True
            WritePN.Visible = False
        ElseIf mode = ItemsCtrl.Mode.Read Then
            ItemIDTB.ReadOnly = True
            SubscribePN.Visible = False
            WritePN.Visible = False
        ElseIf mode = ItemsCtrl.Mode.Write Then
            ItemIDTB.ReadOnly = True
            SubscribePN.Visible = False
            WritePN.Visible = True
        End If

        If item Is Nothing Then
            Exit Function
        End If

        If Not ShowDialog() = DialogResult.OK Then
            Exit Function
        End If

        item.ItemID = ItemIDTB.Text
        item.Active = ActiveCB.Checked
        item.ReqType = DataTypeCTRL.SelectedType
        item.AccessPath = AccessPathTB.Text
        item.Value = ValueCTRL.DataValue

        ShowDialog = True

    End Function

    Public Overloads Function ShowDialog(ByRef pItem As OPCItem) As Boolean

        ShowDialog = False

        ItemIDTB.Text = pItem.ItemID
        ActiveCB.Checked = pItem.IsActive
        DataTypeCTRL.SelectedType = Opc.Type.GetType(pItem.RequestedDataType)
        AccessPathTB.Text = pItem.AccessPath
        ValueCTRL.DataValue = pItem.Value

        ItemIDTB.ReadOnly = False
        SubscribePN.Visible = True
        WritePN.Visible = False
        AccessPathTB.ReadOnly = False

        If Not ShowDialog() = DialogResult.OK Then
            Exit Function
        End If

        Try
            pItem.IsActive = ActiveCB.Checked
            pItem.RequestedDataType = Opc.Type.GetType(DataTypeCTRL.SelectedType)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ShowDialog = True

    End Function
End Class
