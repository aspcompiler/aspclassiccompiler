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
''' A parse tree for a With block statement.
''' </summary>
Public NotInheritable Class WithBlockStatement
    Inherits ExpressionBlockStatement

    ''' <summary>
    ''' Constructs a new parse tree for a With statement block.
    ''' </summary>
    ''' <param name="expression">The expression.</param>
    ''' <param name="statements">The statements in the block.</param>
    ''' <param name="endStatement">The End statement for the block, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal expression As Expression, ByVal statements As StatementCollection, ByVal endStatement As EndBlockStatement, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.WithBlockStatement, expression, statements, endStatement, span, comments)
    End Sub
End Class