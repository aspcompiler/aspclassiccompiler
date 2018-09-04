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
''' A floating point literal.
''' </summary>
Public NotInheritable Class FloatingPointLiteralToken
    Inherits Token

    Private ReadOnly _Literal As Double
    Private ReadOnly _TypeCharacter As TypeCharacter  ' The type character after the literal, if any

    ''' <summary>
    ''' The value of the literal.
    ''' </summary>
    Public ReadOnly Property Literal() As Double
        Get
            Return _Literal
        End Get
    End Property

    ''' <summary>
    ''' The type character after the literal.
    ''' </summary>
    Public ReadOnly Property TypeCharacter() As TypeCharacter
        Get
            Return _TypeCharacter
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new floating point literal token.
    ''' </summary>
    ''' <param name="literal">The literal value.</param>
    ''' <param name="typeCharacter">The type character of the literal.</param>
    ''' <param name="span">The location of the literal.</param>
    Public Sub New(ByVal literal As Double, ByVal typeCharacter As TypeCharacter, ByVal span As Span)
        MyBase.New(TokenType.FloatingPointLiteral, span)

        If typeCharacter <> typeCharacter.None AndAlso _
           typeCharacter <> typeCharacter.SingleSymbol AndAlso typeCharacter <> typeCharacter.SingleChar AndAlso _
           typeCharacter <> typeCharacter.DoubleSymbol AndAlso typeCharacter <> typeCharacter.DoubleChar Then
            Throw New ArgumentOutOfRangeException("typeCharacter")
        End If

        _Literal = literal
        _TypeCharacter = typeCharacter
    End Sub
End Class