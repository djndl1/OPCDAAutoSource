'============================================================================
' TITLE: EnumCtrl.vb
'
' CONTENTS:
' 
' A control used to select a value of an enumeration.
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

Public Class EnumCtrl
    Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents ValuesCB As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ValuesCB = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'ValuesCB
        '
        Me.ValuesCB.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ValuesCB.Location = New System.Drawing.Point(0, 0)
        Me.ValuesCB.Name = "ValuesCB"
        Me.ValuesCB.Size = New System.Drawing.Size(136, 21)
        Me.ValuesCB.TabIndex = 0
        '
        'EnumCtrl
        '
        Me.Controls.Add(Me.ValuesCB)
        Me.Name = "EnumCtrl"
        Me.Size = New System.Drawing.Size(136, 24)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub SetType(ByRef type As System.Type)

        If Not type.IsEnum Then
            Throw New ArgumentException("Not an enumerated data type.", "type")
        End If

        ValuesCB.Items.Clear()

        Dim values As Array = [Enum].GetValues(type)

        For Each value As Object In values
            ValuesCB.Items.Add(value)
        Next

        ValuesCB.Tag = type

    End Sub

    Public Sub SetValue(ByRef value As Object)

        If value Is Nothing Then
            ValuesCB.SelectedIndex = -1
            Exit Sub
        End If

        ValuesCB.SelectedItem = [Enum].ToObject(DirectCast(ValuesCB.Tag, System.Type), value)

    End Sub

    Public Function GetValue() As Object

        If ValuesCB.SelectedIndex = -1 Then
            GetValue = Nothing
            Exit Function
        End If

        GetValue = ValuesCB.SelectedItem

    End Function

End Class
