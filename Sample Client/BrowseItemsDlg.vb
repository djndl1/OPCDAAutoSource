'============================================================================
' TITLE: BrowseItemsDlg.vb
'
' CONTENTS:
' 
' A dialog used to browse the server address space.
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

Public Class BrowseItemsDlg
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
    Friend WithEvents LeftPN As System.Windows.Forms.Panel
    Friend WithEvents SplitterV As System.Windows.Forms.Splitter
    Friend WithEvents BrowseCTRL As VbNetSampleClient.BrowseCtrl
    Friend WithEvents RightPN As System.Windows.Forms.Panel
    Friend WithEvents PropertiesCTRL As VbNetSampleClient.PropertiesCtrl
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.BottomPN = New System.Windows.Forms.Panel
        Me.CancelBTN = New System.Windows.Forms.Button
        Me.OkBTN = New System.Windows.Forms.Button
        Me.LeftPN = New System.Windows.Forms.Panel
        Me.BrowseCTRL = New VbNetSampleClient.BrowseCtrl
        Me.SplitterV = New System.Windows.Forms.Splitter
        Me.RightPN = New System.Windows.Forms.Panel
        Me.PropertiesCTRL = New VbNetSampleClient.PropertiesCtrl
        Me.BottomPN.SuspendLayout()
        Me.LeftPN.SuspendLayout()
        Me.RightPN.SuspendLayout()
        Me.SuspendLayout()
        '
        'BottomPN
        '
        Me.BottomPN.Controls.Add(Me.CancelBTN)
        Me.BottomPN.Controls.Add(Me.OkBTN)
        Me.BottomPN.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.BottomPN.Location = New System.Drawing.Point(0, 334)
        Me.BottomPN.Name = "BottomPN"
        Me.BottomPN.Size = New System.Drawing.Size(1016, 32)
        Me.BottomPN.TabIndex = 1
        '
        'CancelBTN
        '
        Me.CancelBTN.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CancelBTN.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelBTN.Location = New System.Drawing.Point(937, 4)
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
        Me.LeftPN.TabIndex = 2
        '
        'BrowseCTRL
        '
        Me.BrowseCTRL.Dock = System.Windows.Forms.DockStyle.Fill
        Me.BrowseCTRL.Location = New System.Drawing.Point(4, 4)
        Me.BrowseCTRL.Name = "BrowseCTRL"
        Me.BrowseCTRL.Size = New System.Drawing.Size(296, 326)
        Me.BrowseCTRL.TabIndex = 4
        '
        'SplitterV
        '
        Me.SplitterV.Location = New System.Drawing.Point(300, 0)
        Me.SplitterV.Name = "SplitterV"
        Me.SplitterV.Size = New System.Drawing.Size(3, 334)
        Me.SplitterV.TabIndex = 3
        Me.SplitterV.TabStop = False
        '
        'RightPN
        '
        Me.RightPN.Controls.Add(Me.PropertiesCTRL)
        Me.RightPN.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RightPN.DockPadding.Bottom = 4
        Me.RightPN.DockPadding.Right = 4
        Me.RightPN.DockPadding.Top = 4
        Me.RightPN.Location = New System.Drawing.Point(303, 0)
        Me.RightPN.Name = "RightPN"
        Me.RightPN.Size = New System.Drawing.Size(713, 334)
        Me.RightPN.TabIndex = 4
        '
        'PropertiesCTRL
        '
        Me.PropertiesCTRL.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PropertiesCTRL.Location = New System.Drawing.Point(0, 4)
        Me.PropertiesCTRL.Name = "PropertiesCTRL"
        Me.PropertiesCTRL.Size = New System.Drawing.Size(709, 326)
        Me.PropertiesCTRL.TabIndex = 0
        '
        'BrowseItemsDlg
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1016, 366)
        Me.Controls.Add(Me.RightPN)
        Me.Controls.Add(Me.SplitterV)
        Me.Controls.Add(Me.LeftPN)
        Me.Controls.Add(Me.BottomPN)
        Me.Name = "BrowseItemsDlg"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Browse Items"
        Me.BottomPN.ResumeLayout(False)
        Me.LeftPN.ResumeLayout(False)
        Me.RightPN.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private m_pServer As OPCServer

    Public Overloads Function ShowDialog(ByRef pServer As OPCServer) As Boolean

        ShowDialog = False

        If pServer Is Nothing Then
            Exit Function
        End If

        m_pServer = pServer

        'initialize controls
        BrowseCTRL.Connect(m_pServer)
        PropertiesCTRL.Initialize(m_pServer)

        'show dialog
        If Not ShowDialog() = DialogResult.OK Then
            Exit Function
        End If

        ShowDialog = True

    End Function

    Private Sub BrowseCTRL_ItemSelected(ByVal itemID As String) Handles BrowseCTRL.ItemSelected
        PropertiesCTRL.ShowProperties(itemID)
    End Sub
End Class
