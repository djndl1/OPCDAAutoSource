'============================================================================
' TITLE: Server.vb
'
' CONTENTS:
' 
' Defines classes to store client specific information for servers, groups and items.
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

Enum AccessRights
    Readable = 1
    Writable = 2
    ReadWriteable = 3
End Enum

Public Class Item

    Public ItemID As String
    Public ClientHandle As Int32
    Public ServerHandle As Int32
    Public Value As Object
    Public Quality As Int16
    Public Timestamp As Date
    Public Active As Boolean
    Public ReqType As System.Type
    Public AccessPath As String

    Public Shared Function NewItem(ByRef item As OPCItem) As Item

        Dim current As Item = New Item

        current.ItemID = item.ItemID
        current.ClientHandle = item.ClientHandle
        current.ServerHandle = item.ServerHandle
        current.AccessPath = item.AccessPath
        current.ReqType = Opc.Type.GetType(item.RequestedDataType)
        current.Active = item.IsActive
        current.Value = item.Value
        current.Quality = item.Quality
        current.Timestamp = item.TimeStamp

        NewItem = current

    End Function

End Class

Public Class Group
    Public ClientHandle As Int32
    Public ServerHandle As Int32
    Public Items As Hashtable

    Public Sub AddItem(ByRef item As Item)
        If Items Is Nothing Then
            Items = New Hashtable
        End If

        Items.Add(item.ClientHandle, item)
    End Sub

    Public Sub RemoveItem(ByVal ClientHandle As Int32)
        Items.Remove(ClientHandle)
    End Sub

    Public Function AddItem(ByRef TheItem As OPCItem) As Item
        Dim item As Item = item.NewItem(TheItem)
        AddItem(item)
        AddItem = item
    End Function

    Public Function GetItem(ByRef item As OPCItem) As Item
        GetItem = Items.Item(item.ClientHandle)
    End Function

End Class

Public Class Server
    Public Server As OPCServer
    Public Groups As Hashtable
    Private Counter As Int32

    Public Function AddGroup(ByRef pGroup As OPCGroup) As Group

        Dim group As Group = New Group

        group.ServerHandle = pGroup.ServerHandle

        Dim items As OPCItems = pGroup.OPCItems

        For Each item As OPCItem In items
            group.AddItem(item)
        Next

        If Groups Is Nothing Then
            Groups = New Hashtable
        End If

        group.ClientHandle = NextHandle()
        pGroup.ClientHandle = group.ClientHandle

        Groups.Add(group.ClientHandle, group)

        AddGroup = group

    End Function

    Public Sub RemoveGroup(ByVal ClientHandle As Int32)
        Groups.Remove(ClientHandle)
    End Sub

    Public Function GetGroup(ByRef pGroup As OPCGroup) As Group
        GetGroup = Groups.Item(pGroup.ClientHandle)
    End Function

    Public Function NextHandle() As Int32
        Counter = (Counter Mod 10) + Counter + 1
        NextHandle = Counter
    End Function

End Class
