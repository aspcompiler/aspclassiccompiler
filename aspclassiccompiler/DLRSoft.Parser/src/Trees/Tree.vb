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
''' The root class of all trees.
''' </summary>
Public MustInherit Class Tree
    Private ReadOnly _Type As TreeType
    Private ReadOnly _Span As Span
    Private _Parent As Tree
    Private _Children As ReadOnlyCollection(Of Tree)

    ''' <summary>
    ''' The type of the tree.
    ''' </summary>
    Public ReadOnly Property Type() As TreeType
        Get
            Return _Type
        End Get
    End Property

    ''' <summary>
    ''' The location of the tree.
    ''' </summary>
    ''' <remarks>
    ''' The span ends at the first character beyond the tree
    ''' </remarks>
    Public ReadOnly Property Span() As Span
        Get
            Return _Span
        End Get
    End Property

    ''' <summary>
    ''' The parent of the tree. Nothing if the root tree.
    ''' </summary>
    Public ReadOnly Property Parent() As Tree
        Get
            Return _Parent
        End Get
    End Property

    ''' <summary>
    ''' The children of the tree.
    ''' </summary>
    Public ReadOnly Property Children() As ReadOnlyCollection(Of Tree)
        Get
            If _Children Is Nothing Then
                Dim ChildList As List(Of Tree) = New List(Of Tree)

                GetChildTrees(ChildList)
                _Children = New ReadOnlyCollection(Of Tree)(ChildList)
            End If

            Return _Children
        End Get
    End Property

    ''' <summary>
    ''' Whether the tree is 'bad'.
    ''' </summary>
    Public Overridable ReadOnly Property IsBad() As Boolean
        Get
            Return False
        End Get
    End Property

    Protected Sub New(ByVal type As TreeType, ByVal span As Span)
        Debug.Assert(type >= TreeType.SyntaxError AndAlso type <= TreeType.File)
        _Type = type
        _Span = span
    End Sub

    Protected Sub SetParent(ByVal child As Tree)
        If child IsNot Nothing Then
            child._Parent = Me
        End If
    End Sub

    Protected Sub SetParents(Of T As Tree)(ByVal children As IList(Of T))
        If children IsNot Nothing Then
            For Each Child As Tree In children
                SetParent(Child)
            Next
        End If
    End Sub

    Protected Shared Sub AddChild(ByVal childList As IList(Of Tree), ByVal child As Tree)
        If child IsNot Nothing Then
            childList.Add(child)
        End If
    End Sub

    Protected Shared Sub AddChildren(Of T As Tree)(ByVal childList As IList(Of Tree), ByVal children As ReadOnlyCollection(Of T))
        If children IsNot Nothing Then
            For Each Child As Tree In children
                childList.Add(Child)
            Next
        End If
    End Sub

    Protected Overridable Sub GetChildTrees(ByVal childList As IList(Of Tree))
        ' By default, trees have no children
    End Sub
End Class