'============================================================================
' TITLE: MainForm.vb
'
' CONTENTS:
' 
' The main form for the sample client application.
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
Imports System.Runtime.InteropServices

Public Class MainForm
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
    Friend WithEvents GroupsCTRL As VbNetSampleClient.GroupsCtrl
    Friend WithEvents MainMenu As System.Windows.Forms.MainMenu
    Friend WithEvents ServerMI As System.Windows.Forms.MenuItem
    Friend WithEvents ConnectMI As System.Windows.Forms.MenuItem
    Friend WithEvents StatusBar As System.Windows.Forms.StatusBar
    Friend WithEvents LeftPN As System.Windows.Forms.Panel
    Friend WithEvents SplitterV As System.Windows.Forms.Splitter
    Friend WithEvents RightPN As System.Windows.Forms.Panel
    Friend WithEvents GroupUpdatesCTRL As VbNetSampleClient.GroupUpdatesCtrl
    Friend WithEvents DisconnectMI As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupsCTRL = New VbNetSampleClient.GroupsCtrl
        Me.MainMenu = New System.Windows.Forms.MainMenu
        Me.ServerMI = New System.Windows.Forms.MenuItem
        Me.ConnectMI = New System.Windows.Forms.MenuItem
        Me.DisconnectMI = New System.Windows.Forms.MenuItem
        Me.StatusBar = New System.Windows.Forms.StatusBar
        Me.LeftPN = New System.Windows.Forms.Panel
        Me.SplitterV = New System.Windows.Forms.Splitter
        Me.RightPN = New System.Windows.Forms.Panel
        Me.GroupUpdatesCTRL = New VbNetSampleClient.GroupUpdatesCtrl
        Me.LeftPN.SuspendLayout()
        Me.RightPN.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupsCTRL
        '
        Me.GroupsCTRL.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupsCTRL.Location = New System.Drawing.Point(4, 4)
        Me.GroupsCTRL.Name = "GroupsCTRL"
        Me.GroupsCTRL.Size = New System.Drawing.Size(396, 683)
        Me.GroupsCTRL.TabIndex = 3
        '
        'MainMenu
        '
        Me.MainMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.ServerMI})
        '
        'ServerMI
        '
        Me.ServerMI.Index = 0
        Me.ServerMI.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.ConnectMI, Me.DisconnectMI})
        Me.ServerMI.Text = "Server"
        '
        'ConnectMI
        '
        Me.ConnectMI.Index = 0
        Me.ConnectMI.Text = "Connect"
        '
        'DisconnectMI
        '
        Me.DisconnectMI.Index = 1
        Me.DisconnectMI.Text = "Disconnect"
        '
        'StatusBar
        '
        Me.StatusBar.Location = New System.Drawing.Point(0, 691)
        Me.StatusBar.Name = "StatusBar"
        Me.StatusBar.Size = New System.Drawing.Size(1016, 22)
        Me.StatusBar.TabIndex = 4
        '
        'LeftPN
        '
        Me.LeftPN.Controls.Add(Me.GroupsCTRL)
        Me.LeftPN.Dock = System.Windows.Forms.DockStyle.Left
        Me.LeftPN.DockPadding.Bottom = 4
        Me.LeftPN.DockPadding.Left = 4
        Me.LeftPN.DockPadding.Top = 4
        Me.LeftPN.Location = New System.Drawing.Point(0, 0)
        Me.LeftPN.Name = "LeftPN"
        Me.LeftPN.Size = New System.Drawing.Size(400, 691)
        Me.LeftPN.TabIndex = 5
        '
        'SplitterV
        '
        Me.SplitterV.Location = New System.Drawing.Point(400, 0)
        Me.SplitterV.Name = "SplitterV"
        Me.SplitterV.Size = New System.Drawing.Size(3, 691)
        Me.SplitterV.TabIndex = 6
        Me.SplitterV.TabStop = False
        '
        'RightPN
        '
        Me.RightPN.Controls.Add(Me.GroupUpdatesCTRL)
        Me.RightPN.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RightPN.DockPadding.Bottom = 4
        Me.RightPN.DockPadding.Right = 4
        Me.RightPN.DockPadding.Top = 4
        Me.RightPN.Location = New System.Drawing.Point(403, 0)
        Me.RightPN.Name = "RightPN"
        Me.RightPN.Size = New System.Drawing.Size(613, 691)
        Me.RightPN.TabIndex = 7
        '
        'GroupUpdatesCTRL
        '
        Me.GroupUpdatesCTRL.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupUpdatesCTRL.Location = New System.Drawing.Point(0, 4)
        Me.GroupUpdatesCTRL.Name = "GroupUpdatesCTRL"
        Me.GroupUpdatesCTRL.Size = New System.Drawing.Size(609, 683)
        Me.GroupUpdatesCTRL.TabIndex = 0
        '
        'MainForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1016, 713)
        Me.Controls.Add(Me.RightPN)
        Me.Controls.Add(Me.SplitterV)
        Me.Controls.Add(Me.LeftPN)
        Me.Controls.Add(Me.StatusBar)
        Me.Menu = Me.MainMenu
        Me.Name = "MainForm"
        Me.Text = "OPC VB .NET Sample Client"
        Me.LeftPN.ResumeLayout(False)
        Me.RightPN.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private m_aServer As Server

    Private Sub ConnectMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConnectMI.Click

        If Not m_aServer Is Nothing Then

            If Not m_aServer.Server Is Nothing Then
                m_aServer.Server.Disconnect()
                m_aServer.Server = Nothing
            End If

            If Not m_aServer.Groups Is Nothing Then
                m_aServer.Groups.Clear()
            End If

            GroupsCTRL.Initailize(Nothing)
            GroupUpdatesCTRL.Initialize(Nothing)

        End If

        'create a new server object.
        Dim pServer As OPCServer

        Try
            pServer = New OPCServer
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        'select the server.
        Dim node As String
        Dim progID As String = New SelectServerDlg().ShowDialog(pServer, node)

        If progID Is Nothing Then
            Exit Sub
        End If

        'connect to the server.
        Try
            pServer.Connect(progID, node)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        'save the server object.
        m_aServer = New Server
        m_aServer.Server = pServer

        GroupsCTRL.Initailize(m_aServer)
        GroupUpdatesCTRL.Initialize(m_aServer)

    End Sub

    Private Sub DisconnectMI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DisconnectMI.Click

        If Not m_aServer Is Nothing Then

            If Not m_aServer.Server Is Nothing Then
                m_aServer.Server.Disconnect()
                m_aServer.Server = Nothing
            End If

            If Not m_aServer.Groups Is Nothing Then
                m_aServer.Groups.Clear()
            End If

            GroupsCTRL.Initailize(Nothing)
            GroupUpdatesCTRL.Initialize(Nothing)

        End If

    End Sub

    Private Sub GroupsCTRL_GroupSubscribed(ByRef pGroup As OPCGroup) Handles GroupsCTRL.GroupSubscribed
        GroupUpdatesCTRL.GroupSubscribed(pGroup)
    End Sub

End Class


