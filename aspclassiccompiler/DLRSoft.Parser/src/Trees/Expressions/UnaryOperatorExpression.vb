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
''' A parse tree for an unary operator expression.
''' </summary>
Public NotInheritable Class UnaryOperatorExpression
    Inherits UnaryExpression

    Private _Operator As OperatorType

    ''' <summary>
    ''' The operator.
    ''' </summary>
    Public ReadOnly Property [Operator]() As OperatorType
        Get
            Return _Operator
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new unary operator expression parse tree.
    ''' </summary>
    ''' <param name="operator">The type of the unary operator.</param>
    ''' <param name="operand">The operand of the operator.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal [operator] As OperatorType, ByVal operand As Expression, ByVal span As Span)
        MyBase.New(TreeType.UnaryOperatorExpression, operand, span)

        If [operator] < OperatorType.UnaryPlus OrElse [operator] > OperatorType.Not Then
            Throw New ArgumentOutOfRangeException("operator")
        End If

        _Operator = [operator]
    End Sub
End Class