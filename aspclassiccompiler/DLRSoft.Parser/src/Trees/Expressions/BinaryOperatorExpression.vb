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
''' A parse tree for a binary operator expression.
''' </summary>
Public NotInheritable Class BinaryOperatorExpression
    Inherits Expression

    Private ReadOnly _LeftOperand As Expression
    Private ReadOnly _Operator As OperatorType
    Private ReadOnly _OperatorLocation As Location
    Private ReadOnly _RightOperand As Expression

    ''' <summary>
    ''' The left operand expression.
    ''' </summary>
    Public ReadOnly Property LeftOperand() As Expression
        Get
            Return _LeftOperand
        End Get
    End Property

    ''' <summary>
    ''' The operator.
    ''' </summary>
    Public ReadOnly Property [Operator]() As OperatorType
        Get
            Return _Operator
        End Get
    End Property

    ''' <summary>
    ''' The location of the operator.
    ''' </summary>
    Public ReadOnly Property OperatorLocation() As Location
        Get
            Return _OperatorLocation
        End Get
    End Property

    ''' <summary>
    ''' The right operand expression.
    ''' </summary>
    Public ReadOnly Property RightOperand() As Expression
        Get
            Return _RightOperand
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a binary operation.
    ''' </summary>
    ''' <param name="leftOperand">The left operand expression.</param>
    ''' <param name="operator">The operator.</param>
    ''' <param name="operatorLocation">The location of the operator.</param>
    ''' <param name="rightOperand">The right operand expression.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal leftOperand As Expression, ByVal [operator] As OperatorType, ByVal operatorLocation As Location, ByVal rightOperand As Expression, ByVal span As Span)
        MyBase.New(TreeType.BinaryOperatorExpression, span)

        If [operator] < OperatorType.Plus OrElse [operator] > OperatorType.GreaterThanEquals Then
            Throw New ArgumentOutOfRangeException("operator")
        End If

        If leftOperand Is Nothing Then
            Throw New ArgumentNullException("leftOperand")
        End If

        If rightOperand Is Nothing Then
            Throw New ArgumentNullException("rightOperand")
        End If

        SetParent(leftOperand)
        SetParent(rightOperand)

        _LeftOperand = leftOperand
        _Operator = [operator]
        _OperatorLocation = operatorLocation
        _RightOperand = rightOperand
    End Sub

    Public Overrides ReadOnly Property IsConstant() As Boolean
        Get
            Return LeftOperand.IsConstant AndAlso RightOperand.IsConstant
        End Get
    End Property

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, LeftOperand)
        AddChild(childList, RightOperand)
    End Sub
End Class