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
''' A parse tree for an Else If block statement.
''' </summary>
Public NotInheritable Class ElseIfBlockStatement
    Inherits BlockStatement

    Private ReadOnly _ElseIfStatement As ElseIfStatement

    ''' <summary>
    ''' The Else If statement.
    ''' </summary>
    Public ReadOnly Property ElseIfStatement() As ElseIfStatement
        Get
            Return _ElseIfStatement
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for an Else If block statement.
    ''' </summary>
    ''' <param name="elseIfStatement">The Else If statement.</param>
    ''' <param name="statements">The statements in the block.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal elseIfStatement As ElseIfStatement, ByVal statements As StatementCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.ElseIfBlockStatement, statements, span, comments)

        If elseIfStatement Is Nothing Then
            Throw New ArgumentNullException("elseIfStatement")
        End If

        SetParent(elseIfStatement)
        _ElseIfStatement = elseIfStatement
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, ElseIfStatement)
        MyBase.GetChildTrees(childList)
    End Sub
End Class