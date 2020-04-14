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
''' A parse tree for a Finally block statement.
''' </summary>
Public NotInheritable Class FinallyBlockStatement
    Inherits BlockStatement

    Private ReadOnly _FinallyStatement As FinallyStatement

    ''' <summary>
    ''' The Finally statement.
    ''' </summary>
    Public ReadOnly Property FinallyStatement() As FinallyStatement
        Get
            Return _FinallyStatement
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a Finally block statement.
    ''' </summary>
    ''' <param name="finallyStatement">The Finally statement.</param>
    ''' <param name="statements">The statements in the block.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal finallyStatement As FinallyStatement, ByVal statements As StatementCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.FinallyBlockStatement, statements, span, comments)

        If finallyStatement Is Nothing Then
            Throw New ArgumentNullException("finallyStatement")
        End If

        SetParent(finallyStatement)
        _FinallyStatement = finallyStatement
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, FinallyStatement)
        MyBase.GetChildTrees(childList)
    End Sub
End Class