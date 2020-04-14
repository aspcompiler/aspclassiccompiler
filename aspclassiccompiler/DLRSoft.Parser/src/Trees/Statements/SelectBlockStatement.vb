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
''' A parse tree for a Select statement.
''' </summary>
Public NotInheritable Class SelectBlockStatement
    Inherits BlockStatement

    Private ReadOnly _CaseLocation As Location
    Private ReadOnly _Expression As Expression
    Private ReadOnly _CaseBlockStatements As StatementCollection
    Private ReadOnly _CaseElseBlockStatement As CaseElseBlockStatement
    Private ReadOnly _EndStatement As EndBlockStatement

    ''' <summary>
    ''' The location of the 'Case', if any.
    ''' </summary>
    Public ReadOnly Property CaseLocation() As Location
        Get
            Return _CaseLocation
        End Get
    End Property

    ''' <summary>
    ''' The location of the select expression.
    ''' </summary>
    Public ReadOnly Property Expression() As Expression
        Get
            Return _Expression
        End Get
    End Property

    ''' <summary>
    ''' The Case statements.
    ''' </summary>
    Public ReadOnly Property CaseBlockStatements() As StatementCollection
        Get
            Return _CaseBlockStatements
        End Get
    End Property

    ''' <summary>
    ''' The Case Else statement, if any.
    ''' </summary>
    Public ReadOnly Property CaseElseBlockStatement() As CaseElseBlockStatement
        Get
            Return _CaseElseBlockStatement
        End Get
    End Property

    ''' <summary>
    ''' The End Select statement, if any.
    ''' </summary>
    Public ReadOnly Property EndStatement() As EndBlockStatement
        Get
            Return _EndStatement
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a Select statement.
    ''' </summary>
    ''' <param name="caseLocation">The location of the 'Case', if any.</param>
    ''' <param name="expression">The select expression.</param>
    ''' <param name="caseBlockStatements">The Case statements.</param>
    ''' <param name="caseElseBlockStatement">The Case Else statement, if any.</param>
    ''' <param name="endStatement">The End Select statement, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal caseLocation As Location, ByVal expression As Expression, ByVal statements As StatementCollection, ByVal caseBlockStatements As StatementCollection, ByVal caseElseBlockStatement As CaseElseBlockStatement, ByVal endStatement As EndBlockStatement, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.SelectBlockStatement, statements, span, comments)

        If expression Is Nothing Then
            Throw New ArgumentNullException("expression")
        End If

        SetParent(expression)
        SetParent(caseBlockStatements)
        SetParent(caseElseBlockStatement)
        SetParent(endStatement)

        _CaseLocation = caseLocation
        _Expression = expression
        _CaseBlockStatements = caseBlockStatements
        _CaseElseBlockStatement = caseElseBlockStatement
        _EndStatement = endStatement
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Expression)
        AddChild(childList, CaseBlockStatements)
        AddChild(childList, CaseElseBlockStatement)
        AddChild(childList, EndStatement)
    End Sub
End Class