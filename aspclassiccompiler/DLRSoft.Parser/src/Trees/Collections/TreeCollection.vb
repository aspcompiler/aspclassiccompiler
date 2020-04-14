'
' Visual Basic .NET Parser
'
' Copyright (C) 2005, Microsoft Corporation. All rights reserved.
'
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
' MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
'

''' <summary>
''' A collection of a particular type of trees
''' </summary>
''' <typeparam name="T">The type of tree the collection contains.</typeparam>
Public MustInherit Class TreeCollection(Of T As Tree)
    Inherits Tree
    Implements IList(Of T)

    Private _Trees As ReadOnlyCollection(Of T)

    Protected Sub New(ByVal type As TreeType, ByVal trees As IList(Of T), ByVal span As Span)
        MyBase.New(type, span)
        Debug.Assert(type >= TreeType.ArgumentCollection AndAlso type <= TreeType.DeclarationCollection)

        If trees Is Nothing Then
            _Trees = New ReadOnlyCollection(Of T)(New List(Of T)())
        Else
            _Trees = New ReadOnlyCollection(Of T)(trees)
            SetParents(trees)
        End If
    End Sub

    Private Sub Add(ByVal item As T) Implements ICollection(Of T).Add
        Throw New NotSupportedException()
    End Sub

    Private Sub Clear() Implements ICollection(Of T).Clear
        Throw New NotSupportedException()
    End Sub

    Public Function Contains(ByVal item As T) As Boolean Implements ICollection(Of T).Contains
        Return _Trees.Contains(item)
    End Function

    Public Sub CopyTo(ByVal array() As T, ByVal arrayIndex As Integer) Implements ICollection(Of T).CopyTo
        _Trees.CopyTo(array, arrayIndex)
    End Sub

    Public ReadOnly Property Count() As Integer Implements ICollection(Of T).Count
        Get
            Return _Trees.Count
        End Get
    End Property

    Private ReadOnly Property IsReadOnly() As Boolean Implements ICollection(Of T).IsReadOnly
        Get
            Return True
        End Get
    End Property

    Private Function Remove(ByVal item As T) As Boolean Implements ICollection(Of T).Remove
        Throw New NotSupportedException()
    End Function

    Public Function GetEnumerator() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
        Return _Trees.GetEnumerator
    End Function

    Public Function IndexOf(ByVal item As T) As Integer Implements IList(Of T).IndexOf
        Return _Trees.IndexOf(item)
    End Function

    Private Sub Insert(ByVal index As Integer, ByVal item As T) Implements IList(Of T).Insert
        Throw New NotSupportedException()
    End Sub

    Public ReadOnly Property Item(ByVal index As Integer) As T
        Get
            Return _Trees.Item(index)
        End Get
    End Property

    Private Property IListItem(ByVal index As Integer) As T Implements IList(Of T).Item
        Get
            Return _Trees.Item(index)
        End Get
        Set(ByVal value As T)
            Throw New NotSupportedException()
        End Set
    End Property

    Private Sub RemoveAt(ByVal index As Integer) Implements IList(Of T).RemoveAt
        Throw New NotSupportedException()
    End Sub

    Private Function IEnumerableGetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return _Trees.GetEnumerator()
    End Function

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChildren(childList, _Trees)
    End Sub
End Class
