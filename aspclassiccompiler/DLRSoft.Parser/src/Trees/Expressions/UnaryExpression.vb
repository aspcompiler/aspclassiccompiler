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
''' A parse tree for an expression that has an operand.
''' </summary>
Public MustInherit Class UnaryExpression
    Inherits Expression

    Private ReadOnly _Operand As Expression

    ''' <summary>
    ''' The operand of the expression.
    ''' </summary>
    Public ReadOnly Property Operand() As Expression
        Get
            Return _Operand
        End Get
    End Property

    Protected Sub New(ByVal type As TreeType, ByVal operand As Expression, ByVal span As Span)
        MyBase.New(type, span)

        Debug.Assert(type = TreeType.ParentheticalExpression OrElse _
                     type = TreeType.TypeOfExpression OrElse _
                     (type = TreeType.CTypeExpression OrElse type = TreeType.DirectCastExpression OrElse type = TreeType.TryCastExpression) OrElse _
                     type = TreeType.IntrinsicCastExpression OrElse _
                     (type >= TreeType.UnaryOperatorExpression AndAlso type <= TreeType.AddressOfExpression))

        If operand Is Nothing Then
            Throw New ArgumentNullException("operand")
        End If

        SetParent(operand)
        _Operand = operand
    End Sub

    Public Overrides ReadOnly Property IsConstant() As Boolean
        Get
            Return Operand.IsConstant
        End Get
    End Property

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Operand)
    End Sub
End Class