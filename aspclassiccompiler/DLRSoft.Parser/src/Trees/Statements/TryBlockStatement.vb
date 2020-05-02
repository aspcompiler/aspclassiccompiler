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
''' A parse tree for a Try statement.
''' </summary>
Public NotInheritable Class TryBlockStatement
    Inherits BlockStatement

    Private ReadOnly _CatchBlockStatements As StatementCollection
    Private ReadOnly _FinallyBlockStatement As FinallyBlockStatement
    Private ReadOnly _EndStatement As EndBlockStatement

    ''' <summary>
    ''' The Catch statements.
    ''' </summary>
    Public ReadOnly Property CatchBlockStatements() As StatementCollection
        Get
            Return _CatchBlockStatements
        End Get
    End Property

    ''' <summary>
    ''' The Finally statement, if any.
    ''' </summary>
    Public ReadOnly Property FinallyBlockStatement() As FinallyBlockStatement
        Get
            Return _FinallyBlockStatement
        End Get
    End Property

    ''' <summary>
    ''' The End Try statement, if any.
    ''' </summary>
    Public ReadOnly Property EndStatement() As EndBlockStatement
        Get
            Return _EndStatement
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a Try statement.
    ''' </summary>
    ''' <param name="statements">The statements in the Try block.</param>
    ''' <param name="catchBlockStatements">The Catch statements.</param>
    ''' <param name="finallyBlockStatement">The Finally statement, if any.</param>
    ''' <param name="endStatement">The End Try statement, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments of the parse tree.</param>
    Public Sub New(ByVal statements As StatementCollection, ByVal catchBlockStatements As StatementCollection, ByVal finallyBlockStatement As FinallyBlockStatement, ByVal endStatement As EndBlockStatement, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.TryBlockStatement, statements, span, comments)

        SetParent(catchBlockStatements)
        SetParent(finallyBlockStatement)
        SetParent(endStatement)

        _CatchBlockStatements = catchBlockStatements
        _FinallyBlockStatement = finallyBlockStatement
        _EndStatement = endStatement
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)
        AddChild(childList, CatchBlockStatements)
        AddChild(childList, FinallyBlockStatement)
        AddChild(childList, EndStatement)
    End Sub
End Class