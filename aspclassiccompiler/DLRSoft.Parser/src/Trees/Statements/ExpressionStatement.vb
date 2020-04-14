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
''' A parse tree for an expression statement.
''' </summary>
Public MustInherit Class ExpressionStatement
    Inherits Statement

    Private ReadOnly _Expression As Expression

    ''' <summary>
    ''' The expression.
    ''' </summary>
    Public ReadOnly Property Expression() As Expression
        Get
            Return _Expression
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for an expression statement.
    ''' </summary>
    ''' <param name="type">The type of the parse tree.</param>
    ''' <param name="expression">The expression.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Protected Sub New(ByVal type As TreeType, ByVal expression As Expression, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(type, span, comments)

        Debug.Assert(type = TreeType.ReturnStatement OrElse type = TreeType.ErrorStatement OrElse type = TreeType.ThrowStatement)

        SetParent(expression)
        _Expression = expression
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Expression)
    End Sub
End Class