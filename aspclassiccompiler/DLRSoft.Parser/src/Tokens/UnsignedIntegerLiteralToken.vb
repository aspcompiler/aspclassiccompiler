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
''' An integer literal.
''' </summary>
Public NotInheritable Class UnsignedIntegerLiteralToken
    Inherits Token

    Private ReadOnly _Literal As ULong
    Private ReadOnly _TypeCharacter As TypeCharacter  ' The type character after the literal, if any
    Private ReadOnly _IntegerBase As IntegerBase      ' The base of the literal

    ''' <summary>
    ''' The value of the literal.
    ''' </summary>
    <CLSCompliant(False)> _
    Public ReadOnly Property Literal() As ULong
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
    ''' The integer base of the literal.
    ''' </summary>
    Public ReadOnly Property IntegerBase() As IntegerBase
        Get
            Return _IntegerBase
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new unsigned integer literal.
    ''' </summary>
    ''' <param name="literal">The literal value.</param>
    ''' <param name="integerBase">The integer base of the literal.</param>
    ''' <param name="typeCharacter">The type character of the literal.</param>
    ''' <param name="span">The location of the literal.</param>
    <CLSCompliant(False)> _
    Public Sub New(ByVal literal As ULong, ByVal integerBase As IntegerBase, ByVal typeCharacter As TypeCharacter, ByVal span As Span)
        MyBase.New(TokenType.UnsignedIntegerLiteral, span)

        If integerBase < integerBase.Decimal OrElse integerBase > integerBase.Hexadecimal Then
            Throw New ArgumentOutOfRangeException("integerBase")
        End If

        If typeCharacter <> typeCharacter.None AndAlso _
           typeCharacter <> typeCharacter.UnsignedIntegerChar AndAlso _
           typeCharacter <> typeCharacter.UnsignedLongChar AndAlso _
           typeCharacter <> typeCharacter.UnsignedShortChar Then
            Throw New ArgumentOutOfRangeException("typeCharacter")
        End If

        _Literal = literal
        _IntegerBase = integerBase
        _TypeCharacter = typeCharacter
    End Sub
End Class

