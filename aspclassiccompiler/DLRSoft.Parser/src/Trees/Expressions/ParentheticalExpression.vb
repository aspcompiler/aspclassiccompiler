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
''' A parse tree for a parenthesized expression.
''' </summary>
Public NotInheritable Class ParentheticalExpression
    Inherits UnaryExpression

    Private ReadOnly _RightParenthesisLocation As Location

    ''' <summary>
    ''' The location of the ')'.
    ''' </summary>
    Public ReadOnly Property RightParenthesisLocation() As Location
        Get
            Return _RightParenthesisLocation
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parenthesized expression parse tree.
    ''' </summary>
    ''' <param name="operand">The operand of the expression.</param>
    ''' <param name="rightParenthesisLocation">The location of the ')'.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal operand As Expression, ByVal rightParenthesisLocation As Location, ByVal span As Span)
        MyBase.New(TreeType.ParentheticalExpression, operand, span)

        _RightParenthesisLocation = rightParenthesisLocation
    End Sub
End Class