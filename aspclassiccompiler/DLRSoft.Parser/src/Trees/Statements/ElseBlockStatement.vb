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
''' A parse tree for an Else block statement.
''' </summary>
Public NotInheritable Class ElseBlockStatement
    Inherits BlockStatement

    Private ReadOnly _ElseStatement As ElseStatement

    ''' <summary>
    ''' The Else or Else If statement.
    ''' </summary>
    Public ReadOnly Property ElseStatement() As ElseStatement
        Get
            Return _ElseStatement
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for an Else block statement.
    ''' </summary>
    ''' <param name="elseStatement">The Else statement.</param>
    ''' <param name="statements">The statements in the block.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal elseStatement As ElseStatement, ByVal statements As StatementCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.ElseBlockStatement, statements, span, comments)

        If elseStatement Is Nothing Then
            Throw New ArgumentNullException("elseStatement")
        End If

        SetParent(elseStatement)
        _ElseStatement = elseStatement
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, ElseStatement)
        MyBase.GetChildTrees(childList)
    End Sub
End Class