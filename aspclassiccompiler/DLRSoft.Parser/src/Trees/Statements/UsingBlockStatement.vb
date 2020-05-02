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
''' A parse tree for a Using block statement.
''' </summary>
Public NotInheritable Class UsingBlockStatement
    Inherits ExpressionBlockStatement

    Private ReadOnly _VariableDeclarators As VariableDeclaratorCollection

    ''' <summary>
    ''' The variable declarators, if no expression.
    ''' </summary>
    Public ReadOnly Property VariableDeclarators() As VariableDeclaratorCollection
        Get
            Return _VariableDeclarators
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a Using statement block with an expression.
    ''' </summary>
    ''' <param name="expression">The expression.</param>
    ''' <param name="statements">The statements in the block.</param>
    ''' <param name="endStatement">The End statement for the block, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal expression As Expression, ByVal statements As StatementCollection, ByVal endStatement As EndBlockStatement, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.UsingBlockStatement, expression, statements, endStatement, span, comments)
    End Sub

    ''' <summary>
    ''' Constructs a new parse tree for a Using statement block with variable declarators.
    ''' </summary>
    ''' <param name="variableDeclarators">The variable declarators.</param>
    ''' <param name="statements">The statements in the block.</param>
    ''' <param name="endStatement">The End statement for the block, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal variableDeclarators As VariableDeclaratorCollection, ByVal statements As StatementCollection, ByVal endStatement As EndBlockStatement, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.UsingBlockStatement, Nothing, statements, endStatement, span, comments)

        If variableDeclarators Is Nothing Then
            Throw New ArgumentNullException("variableDeclarators")
        End If

        SetParent(variableDeclarators)

        _VariableDeclarators = variableDeclarators
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, VariableDeclarators)
        MyBase.GetChildTrees(childList)
    End Sub
End Class