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
''' A parse tree for a Do statement.
''' </summary>
Public NotInheritable Class DoBlockStatement
    Inherits BlockStatement

    Private ReadOnly _IsWhile As Boolean
    Private ReadOnly _WhileOrUntilLocation As Location
    Private ReadOnly _Expression As Expression
    Private ReadOnly _EndStatement As LoopStatement

    ''' <summary>
    ''' Whether the Do is followed by a While or Until, if any.
    ''' </summary>
    Public ReadOnly Property IsWhile() As Boolean
        Get
            Return _IsWhile
        End Get
    End Property

    ''' <summary>
    ''' The location of the While or Until, if any.
    ''' </summary>
    Public ReadOnly Property WhileOrUntilLocation() As Location
        Get
            Return _WhileOrUntilLocation
        End Get
    End Property

    ''' <summary>
    ''' The While or Until expression, if any.
    ''' </summary>
    Public ReadOnly Property Expression() As Expression
        Get
            Return _Expression
        End Get
    End Property

    ''' <summary>
    ''' The ending Loop statement.
    ''' </summary>
    Public ReadOnly Property EndStatement() As LoopStatement
        Get
            Return _EndStatement
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a Do statement.
    ''' </summary>
    ''' <param name="expression">The While or Until expression, if any.</param>
    ''' <param name="isWhile">Whether the Do is followed by a While or Until, if any.</param>
    ''' <param name="whileOrUntilLocation">The location of the While or Until, if any.</param>
    ''' <param name="statements">The statements in the block.</param>
    ''' <param name="endStatement">The ending Loop statement.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments on the parse tree.</param>
    Public Sub New(ByVal expression As Expression, ByVal isWhile As Boolean, ByVal whileOrUntilLocation As Location, ByVal statements As StatementCollection, ByVal endStatement As LoopStatement, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.DoBlockStatement, statements, span, comments)

        SetParent(expression)
        SetParent(endStatement)

        _Expression = expression
        _IsWhile = isWhile
        _WhileOrUntilLocation = whileOrUntilLocation
        _EndStatement = endStatement
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Expression)
        MyBase.GetChildTrees(childList)
        AddChild(childList, EndStatement)
    End Sub
End Class