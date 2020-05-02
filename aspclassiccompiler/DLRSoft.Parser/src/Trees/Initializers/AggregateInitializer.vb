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
''' A parse tree for an aggregate initializer.
''' </summary>
Public NotInheritable Class AggregateInitializer
    Inherits Initializer

    Private ReadOnly _Elements As InitializerCollection

    ''' <summary>
    ''' The elements of the aggregate initializer.
    ''' </summary>
    Public ReadOnly Property Elements() As InitializerCollection
        Get
            Return _Elements
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new aggregate initializer parse tree.
    ''' </summary>
    ''' <param name="elements">The elements of the aggregate initializer.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal elements As InitializerCollection, ByVal span As Span)
        MyBase.New(TreeType.AggregateInitializer, span)

        If elements Is Nothing Then
            Throw New ArgumentNullException("elements")
        End If

        SetParent(elements)
        _Elements = elements
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Elements)
    End Sub
End Class