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
''' A parse tree for a line If statement.
''' </summary>
Public NotInheritable Class LineIfStatement
    Inherits Statement

    Private ReadOnly _Expression As Expression
    Private ReadOnly _ThenLocation As Location
    Private ReadOnly _IfStatements As StatementCollection
    Private ReadOnly _ElseLocation As Location
    Private ReadOnly _ElseStatements As StatementCollection

    ''' <summary>
    ''' The conditional expression.
    ''' </summary>
    Public ReadOnly Property Expression() As Expression
        Get
            Return _Expression
        End Get
    End Property

    ''' <summary>
    ''' The location of the 'Then'.
    ''' </summary>
    Public ReadOnly Property ThenLocation() As Location
        Get
            Return _ThenLocation
        End Get
    End Property

    ''' <summary>
    ''' The If statements.
    ''' </summary>
    Public ReadOnly Property IfStatements() As StatementCollection
        Get
            Return _IfStatements
        End Get
    End Property

    ''' <summary>
    ''' The location of the 'Else', if any.
    ''' </summary>
    Public ReadOnly Property ElseLocation() As Location
        Get
            Return _ElseLocation
        End Get
    End Property

    ''' <summary>
    ''' The Else statements.
    ''' </summary>
    Public ReadOnly Property ElseStatements() As StatementCollection
        Get
            Return _ElseStatements
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a line If statement.
    ''' </summary>
    ''' <param name="expression">The conditional expression.</param>
    ''' <param name="thenLocation">The location of the 'Then'.</param>
    ''' <param name="ifStatements">The If statements.</param>
    ''' <param name="elseLocation">The location of the 'Else', if any.</param>
    ''' <param name="elseStatements">The Else statements.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal expression As Expression, ByVal thenLocation As Location, ByVal ifStatements As StatementCollection, ByVal elseLocation As Location, ByVal elseStatements As StatementCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.LineIfBlockStatement, span, comments)

        If expression Is Nothing Then
            Throw New ArgumentNullException("expression")
        End If

        SetParent(expression)
        SetParent(ifStatements)
        SetParent(elseStatements)

        _Expression = expression
        _ThenLocation = thenLocation
        _IfStatements = ifStatements
        _ElseLocation = elseLocation
        _ElseStatements = elseStatements
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Expression)
        AddChild(childList, IfStatements)
        AddChild(childList, ElseStatements)
    End Sub
End Class