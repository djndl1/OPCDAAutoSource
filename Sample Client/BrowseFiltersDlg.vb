'============================================================================
' TITLE: BrowseFiltersDlg.vb
'
' CONTENTS:
' 
' A dialog used to set any filters to use when browsing the server address space.
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

Public Class BrowseFiltersDlg
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        DataTypeCTRL.SetType(GetType(DataType))
        AccessRightsCTRL.SetType(GetType(AccessRights))

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
    Friend WithEvents OkBTN As System.Windows.Forms.Button
    Friend WithEvents CancelBTN As System.Windows.Forms.Button
    Friend WithEvents DataTypeLB As System.Windows.Forms.Label
    Friend WithEvents FilterLB As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents FilterTB As System.Windows.Forms.TextBox
    Friend WithEvents DataTypeCTRL As VbNetSampleClient.EnumCtrl
    Friend WithEvents AccessRightsCTRL As VbNetSampleClient.EnumCtrl
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.BottomPN = New System.Windows.Forms.Panel
        Me.CancelBTN = New System.Windows.Forms.Button
        Me.OkBTN = New System.Windows.Forms.Button
        Me.DataTypeLB = New System.Windows.Forms.Label
        Me.FilterLB = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.FilterTB = New System.Windows.Forms.TextBox
        Me.DataTypeCTRL = New VbNetSampleClient.EnumCtrl
        Me.AccessRightsCTRL = New VbNetSampleClient.EnumCtrl
        Me.BottomPN.SuspendLayout()
        Me.SuspendLayout()
        '
        'BottomPN
        '
        Me.BottomPN.Controls.Add(Me.CancelBTN)
        Me.BottomPN.Controls.Add(Me.OkBTN)
        Me.BottomPN.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.BottomPN.Location = New System.Drawing.Point(0, 86)
        Me.BottomPN.Name = "BottomPN"
        Me.BottomPN.Size = New System.Drawing.Size(280, 32)
        Me.BottomPN.TabIndex = 0
        '
        'CancelBTN
        '
        Me.CancelBTN.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CancelBTN.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelBTN.Location = New System.Drawing.Point(201, 4)
        Me.CancelBTN.Name = "CancelBTN"
        Me.CancelBTN.TabIndex = 1
        Me.CancelBTN.Text = "Cancel"
        '
        'OkBTN
        '
        Me.OkBTN.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OkBTN.Location = New System.Drawing.Point(4, 4)
        Me.OkBTN.Name = "OkBTN"
        Me.OkBTN.TabIndex = 0
        Me.OkBTN.Text = "OK"
        '
        'DataTypeLB
        '
        Me.DataTypeLB.Location = New System.Drawing.Point(4, 32)
        Me.DataTypeLB.Name = "DataTypeLB"
        Me.DataTypeLB.Size = New System.Drawing.Size(76, 23)
        Me.DataTypeLB.TabIndex = 1
        Me.DataTypeLB.Text = "Data Type"
        Me.DataTypeLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FilterLB
        '
        Me.FilterLB.Location = New System.Drawing.Point(4, 4)
        Me.FilterLB.Name = "FilterLB"
        Me.FilterLB.Size = New System.Drawing.Size(76, 23)
        Me.FilterLB.TabIndex = 2
        Me.FilterLB.Text = "Filter"
        Me.FilterLB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(4, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 23)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Access Rights"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FilterTB
        '
        Me.FilterTB.Location = New System.Drawing.Point(80, 8)
        Me.FilterTB.Name = "FilterTB"
        Me.FilterTB.Size = New System.Drawing.Size(192, 20)
        Me.FilterTB.TabIndex = 4
        Me.FilterTB.Text = ""
        '
        'DataTypeCTRL
        '
        Me.DataTypeCTRL.Location = New System.Drawing.Point(80, 32)
        Me.DataTypeCTRL.Name = "DataTypeCTRL"
        Me.DataTypeCTRL.Size = New System.Drawing.Size(136, 24)
        Me.DataTypeCTRL.TabIndex = 5
        '
        'AccessRightsCTRL
        '
        Me.AccessRightsCTRL.Location = New System.Drawing.Point(80, 56)
        Me.AccessRightsCTRL.Name = "AccessRightsCTRL"
        Me.AccessRightsCTRL.Size = New System.Drawing.Size(136, 24)
        Me.AccessRightsCTRL.TabIndex = 6
        '
        'BrowseFiltersDlg
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.CancelBTN
        Me.ClientSize = New System.Drawing.Size(280, 118)
        Me.Controls.Add(Me.AccessRightsCTRL)
        Me.Controls.Add(Me.DataTypeCTRL)
        Me.Controls.Add(Me.FilterTB)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.FilterLB)
        Me.Controls.Add(Me.DataTypeLB)
        Me.Controls.Add(Me.BottomPN)
        Me.Name = "BrowseFiltersDlg"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Browse Filters"
        Me.BottomPN.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Enum DataType
        VT_EMPTY = 0
        VT_NULL = 1
        VT_I2 = 2
        VT_I4 = 3
        VT_R4 = 4
        VT_R8 = 5
        VT_CY = 6
        VT_DATE = 7
        VT_BSTR = 8
        VT_DISPATCH = 9
        VT_ERROR = 10
        VT_BOOL = 11
        VT_VARIANT = 12
        VT_UNKNOWN = 13
        VT_DECIMAL = 14
        VT_I1 = 16
        VT_UI1 = 17
        VT_UI2 = 18
        VT_UI4 = 19
    End Enum

    Public Overloads Function ShowDialog(ByRef browser As OPCBrowser) As Boolean

        ShowDialog = False

        If browser Is Nothing Then
            Exit Function
        End If

        FilterTB.Text = browser.Filter
        DataTypeCTRL.SetValue(browser.DataType)
        AccessRightsCTRL.SetValue(browser.AccessRights)

        If Not ShowDialog() = DialogResult.OK Then
            Exit Function
        End If

        browser.Filter = FilterTB.Text
        browser.DataType = DataTypeCTRL.GetValue()
        browser.AccessRights = AccessRightsCTRL.GetValue()

        ShowDialog = True

    End Function

End Class
