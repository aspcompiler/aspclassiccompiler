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
''' A parse tree for a literal expression.
''' </summary>
Public MustInherit Class LiteralExpression
    Inherits Expression

    Public NotOverridable Overrides ReadOnly Property IsConstant() As Boolean
        Get
            Return True
        End Get
    End Property

    'LC add property to get value
    Public MustOverride ReadOnly Property Value() As Object

    Protected Sub New(ByVal type As TreeType, ByVal span As Span)
        MyBase.New(type, span)

        Debug.Assert(type >= TreeType.StringLiteralExpression AndAlso type <= TreeType.BooleanLiteralExpression)
    End Sub
End Class