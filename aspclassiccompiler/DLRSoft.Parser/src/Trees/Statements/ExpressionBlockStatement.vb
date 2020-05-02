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
''' A parse tree for an expression block statement.
''' </summary>
Public MustInherit Class ExpressionBlockStatement
    Inherits BlockStatement

    Private ReadOnly _Expression As Expression
    Private ReadOnly _EndStatement As EndBlockStatement

    ''' <summary>
    ''' The expression.
    ''' </summary>
    Public ReadOnly Property Expression() As Expression
        Get
            Return _Expression
        End Get
    End Property

    ''' <summary>
    ''' The End statement for the block, if any.
    ''' </summary>
    Public ReadOnly Property EndStatement() As EndBlockStatement
        Get
            Return _EndStatement
        End Get
    End Property

    Protected Sub New(ByVal type As TreeType, ByVal expression As Expression, ByVal statements As StatementCollection, ByVal endStatement As EndBlockStatement, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(type, statements, span, comments)

        Debug.Assert(type = TreeType.WithBlockStatement OrElse _
                     type = TreeType.SyncLockBlockStatement OrElse _
                     type = TreeType.WhileBlockStatement OrElse _
                     type = TreeType.UsingBlockStatement)

        SetParent(expression)
        SetParent(endStatement)

        _Expression = expression
        _EndStatement = endStatement
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Expression)
        MyBase.GetChildTrees(childList)
        AddChild(childList, EndStatement)
    End Sub
End Class