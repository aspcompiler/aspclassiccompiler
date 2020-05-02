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
''' A parse tree for an intrinsic conversion expression.
''' </summary>
Public NotInheritable Class IntrinsicCastExpression
    Inherits UnaryExpression

    Private ReadOnly _IntrinsicType As IntrinsicType
    Private ReadOnly _LeftParenthesisLocation As Location
    Private ReadOnly _RightParenthesisLocation As Location

    ''' <summary>
    ''' The intrinsic type conversion.
    ''' </summary>
    Public ReadOnly Property IntrinsicType() As IntrinsicType
        Get
            Return _IntrinsicType
        End Get
    End Property

    ''' <summary>
    ''' The location of the '('.
    ''' </summary>
    Public ReadOnly Property LeftParenthesisLocation() As Location
        Get
            Return _LeftParenthesisLocation
        End Get
    End Property

    ''' <summary>
    ''' The location of the ')'.
    ''' </summary>
    Public ReadOnly Property RightParenthesisLocation() As Location
        Get
            Return _RightParenthesisLocation
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for an intrinsic conversion expression.
    ''' </summary>
    ''' <param name="intrinsicType">The intrinsic type conversion.</param>
    ''' <param name="leftParenthesisLocation">The location of the '('.</param>
    ''' <param name="operand">The expression to convert.</param>
    ''' <param name="rightParenthesisLocation">The location of the ')'.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal intrinsicType As IntrinsicType, ByVal leftParenthesisLocation As Location, ByVal operand As Expression, ByVal rightParenthesisLocation As Location, ByVal span As Span)
        MyBase.New(TreeType.IntrinsicCastExpression, operand, span)

        If intrinsicType < intrinsicType.Boolean OrElse intrinsicType > intrinsicType.Object Then
            Throw New ArgumentOutOfRangeException("intrinsicType")
        End If

        If operand Is Nothing Then
            Throw New ArgumentNullException("operand")
        End If

        _IntrinsicType = intrinsicType
        _LeftParenthesisLocation = leftParenthesisLocation
        _RightParenthesisLocation = rightParenthesisLocation
    End Sub
End Class