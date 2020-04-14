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
''' A parse tree for an If block.
''' </summary>
Public NotInheritable Class IfBlockStatement
    Inherits BlockStatement

    Private ReadOnly _Expression As Expression
    Private ReadOnly _ThenLocation As Location
    Private ReadOnly _ElseIfBlockStatements As StatementCollection
    Private ReadOnly _ElseBlockStatement As ElseBlockStatement
    Private ReadOnly _EndStatement As EndBlockStatement

    ''' <summary>
    ''' The conditional expression.
    ''' </summary>
    Public ReadOnly Property Expression() As Expression
        Get
            Return _Expression
        End Get
    End Property

    ''' <summary>
    ''' The location of the 'Then', if any.
    ''' </summary>
    Public ReadOnly Property ThenLocation() As Location
        Get
            Return _ThenLocation
        End Get
    End Property

    ''' <summary>
    ''' The Else If statements.
    ''' </summary>
    Public ReadOnly Property ElseIfBlockStatements() As StatementCollection
        Get
            Return _ElseIfBlockStatements
        End Get
    End Property

    ''' <summary>
    ''' The Else statement, if any.
    ''' </summary>
    Public ReadOnly Property ElseBlockStatement() As ElseBlockStatement
        Get
            Return _ElseBlockStatement
        End Get
    End Property

    ''' <summary>
    ''' The End If statement, if any.
    ''' </summary>
    Public ReadOnly Property EndStatement() As EndBlockStatement
        Get
            Return _EndStatement
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a If statement.
    ''' </summary>
    ''' <param name="expression">The conditional expression.</param>
    ''' <param name="thenLocation">The location of the 'Then', if any.</param>
    ''' <param name="statements">The statements in the If block.</param>
    ''' <param name="elseIfBlockStatements">The Else If statements.</param>
    ''' <param name="elseBlockStatement">The Else statement, if any.</param>
    ''' <param name="endStatement">The End If statement, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal expression As Expression, ByVal thenLocation As Location, ByVal statements As StatementCollection, ByVal elseIfBlockStatements As StatementCollection, ByVal elseBlockStatement As ElseBlockStatement, ByVal endStatement As EndBlockStatement, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.IfBlockStatement, statements, span, comments)

        If expression Is Nothing Then
            Throw New ArgumentNullException("expression")
        End If

        SetParent(expression)
        SetParent(elseIfBlockStatements)
        SetParent(elseBlockStatement)
        SetParent(endStatement)

        _Expression = expression
        _ThenLocation = thenLocation
        _ElseIfBlockStatements = elseIfBlockStatements
        _ElseBlockStatement = elseBlockStatement
        _EndStatement = endStatement
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Expression)
        MyBase.GetChildTrees(childList)
        AddChild(childList, ElseIfBlockStatements)
        AddChild(childList, ElseBlockStatement)
        AddChild(childList, EndStatement)
    End Sub
End Class