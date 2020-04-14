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
''' A parse tree for an Error statement.
''' </summary>
Public NotInheritable Class ErrorStatement
    Inherits ExpressionStatement

    ''' <summary>
    ''' Constructs a new parse tree for an Error statement.
    ''' </summary>
    ''' <param name="expression">The expression.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal expression As Expression, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.ErrorStatement, expression, span, comments)
    End Sub
End Class