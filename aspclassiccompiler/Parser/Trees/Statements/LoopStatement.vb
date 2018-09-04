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
''' A parse tree for a Loop statement.
''' </summary>
Public NotInheritable Class LoopStatement
    Inherits Statement

    Private ReadOnly _IsWhile As Boolean
    Private ReadOnly _WhileOrUntilLocation As Location
    Private ReadOnly _Expression As Expression

    ''' <summary>
    ''' Whether the Loop has a While or Until.
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
    ''' The loop expression, if any.
    ''' </summary>
    Public ReadOnly Property Expression() As Expression
        Get
            Return _Expression
        End Get
    End Property

    ''' <summary>
    ''' Constructs a parse tree for a Loop statement.
    ''' </summary>
    ''' <param name="expression">The loop expression, if any.</param>
    ''' <param name="isWhile">WHether the Loop has a While or Until.</param>
    ''' <param name="whileOrUntilLocation">The location of the While or Until, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal expression As Expression, ByVal isWhile As Boolean, ByVal whileOrUntilLocation As Location, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.LoopStatement, span, comments)

        SetParent(expression)

        _Expression = expression
        _IsWhile = isWhile
        _WhileOrUntilLocation = whileOrUntilLocation
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Expression)
    End Sub
End Class