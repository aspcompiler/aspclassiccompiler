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
''' A date/time literal.
''' </summary>
Public NotInheritable Class DateLiteralToken
    Inherits Token

    Private ReadOnly _Literal As Date

    ''' <summary>
    ''' The literal value.
    ''' </summary>
    Public ReadOnly Property Literal() As Date
        Get
            Return _Literal
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new date literal instance.
    ''' </summary>
    ''' <param name="literal">The literal value.</param>
    ''' <param name="span">The location of the literal.</param>
    Public Sub New(ByVal literal As Date, ByVal span As Span)
        MyBase.New(TokenType.DateLiteral, span)
        _Literal = literal
    End Sub
End Class