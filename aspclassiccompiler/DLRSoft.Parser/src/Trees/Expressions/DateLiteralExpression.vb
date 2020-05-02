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
''' A parse tree for a date literal expression.
''' </summary>
Public NotInheritable Class DateLiteralExpression
    Inherits LiteralExpression

    Private ReadOnly _Literal As Date

    ''' <summary>
    ''' The literal value.
    ''' </summary>
    Public ReadOnly Property Literal() As Date
        Get
            Return _Literal
        End Get
    End Property

    'LC
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return _Literal
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a date literal.
    ''' </summary>
    ''' <param name="literal">The literal value.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal literal As Date, ByVal span As Span)
        MyBase.New(TreeType.DateLiteralExpression, span)
        _Literal = literal
    End Sub
End Class