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
''' A parse tree for an expression.
''' </summary>
Public Class Expression
    Inherits Tree

    ''' <summary>
    ''' Creates a bad expression.
    ''' </summary>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <returns>A bad expression.</returns>
    Public Shared Function GetBadExpression(ByVal span As Span) As Expression
        Return New Expression(span)
    End Function

    ''' <summary>
    ''' Whether the expression is constant or not.
    ''' </summary>
    Public Overridable ReadOnly Property IsConstant() As Boolean
        Get
            Return False
        End Get
    End Property

    Protected Sub New(ByVal type As TreeType, ByVal span As Span)
        MyBase.New(type, span)

        Debug.Assert(type >= TreeType.SimpleNameExpression AndAlso type <= TreeType.GetTypeExpression)
    End Sub

    Private Sub New(ByVal span As Span)
        MyBase.New(TreeType.SyntaxError, span)
    End Sub

    Public Overrides ReadOnly Property IsBad() As Boolean
        Get
            Return Type = TreeType.SyntaxError
        End Get
    End Property
End Class