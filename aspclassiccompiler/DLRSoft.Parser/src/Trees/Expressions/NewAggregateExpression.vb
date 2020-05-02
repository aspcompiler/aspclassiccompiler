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
''' A parse tree for a New array expression.
''' </summary>
Public NotInheritable Class NewAggregateExpression
    Inherits Expression

    Private ReadOnly _Target As ArrayTypeName
    Private ReadOnly _Initializer As AggregateInitializer

    ''' <summary>
    ''' The target array type to create.
    ''' </summary>
    Public ReadOnly Property Target() As ArrayTypeName
        Get
            Return _Target
        End Get
    End Property

    ''' <summary>
    ''' The initializer for the array.
    ''' </summary>
    Public ReadOnly Property Initializer() As AggregateInitializer
        Get
            Return _Initializer
        End Get
    End Property

    ''' <summary>
    ''' The constructor for a New array expression parse tree.
    ''' </summary>
    ''' <param name="target">The target array type to create.</param>
    ''' <param name="initializer">The initializer for the array.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal target As ArrayTypeName, ByVal initializer As AggregateInitializer, ByVal span As Span)
        MyBase.New(TreeType.NewAggregateExpression, span)

        If target Is Nothing Then
            Throw New ArgumentNullException("target")
        End If

        If initializer Is Nothing Then
            Throw New ArgumentNullException("initializer")
        End If

        SetParent(target)
        SetParent(initializer)

        _Target = target
        _Initializer = initializer
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Target)
        AddChild(childList, Initializer)
    End Sub
End Class