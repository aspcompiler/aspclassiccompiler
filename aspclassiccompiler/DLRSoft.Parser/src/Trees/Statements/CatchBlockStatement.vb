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
''' A parse tree for a Catch block statement.
''' </summary>
Public NotInheritable Class CatchBlockStatement
    Inherits BlockStatement

    Private ReadOnly _CatchStatement As CatchStatement

    ''' <summary>
    ''' The Catch statement.
    ''' </summary>
    Public ReadOnly Property CatchStatement() As CatchStatement
        Get
            Return _CatchStatement
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a Catch block statement.
    ''' </summary>
    ''' <param name="catchStatement">The Catch statement.</param>
    ''' <param name="statements">The statements in the block.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal catchStatement As CatchStatement, ByVal statements As StatementCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.CatchBlockStatement, statements, span, comments)

        If catchStatement Is Nothing Then
            Throw New ArgumentNullException("catchStatement")
        End If

        SetParent(catchStatement)
        _CatchStatement = catchStatement
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, CatchStatement)
        MyBase.GetChildTrees(childList)
    End Sub
End Class