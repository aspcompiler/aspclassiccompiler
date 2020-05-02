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
''' A decimal literal token.
''' </summary>
Public NotInheritable Class DecimalLiteralToken
    Inherits Token

    Private ReadOnly _Literal As Decimal
    Private ReadOnly _TypeCharacter As TypeCharacter  ' The type character after the literal, if any

    ''' <summary>
    ''' The literal value.
    ''' </summary>
    Public ReadOnly Property Literal() As Decimal
        Get
            Return _Literal
        End Get
    End Property

    ''' <summary>
    ''' The type character of the literal.
    ''' </summary>
    Public ReadOnly Property TypeCharacter() As TypeCharacter
        Get
            Return _TypeCharacter
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new decimal literal token.
    ''' </summary>
    ''' <param name="literal">The literal value.</param>
    ''' <param name="typeCharacter">The literal's type character.</param>
    ''' <param name="span">The location of the literal.</param>
    Public Sub New(ByVal literal As Decimal, ByVal typeCharacter As TypeCharacter, ByVal span As Span)
        MyBase.New(TokenType.DecimalLiteral, span)

        If typeCharacter <> typeCharacter.None AndAlso typeCharacter <> typeCharacter.DecimalChar AndAlso _
           typeCharacter <> typeCharacter.DecimalSymbol Then
            Throw New ArgumentOutOfRangeException("typeCharacter")
        End If

        _Literal = literal
        _TypeCharacter = typeCharacter
    End Sub
End Class
