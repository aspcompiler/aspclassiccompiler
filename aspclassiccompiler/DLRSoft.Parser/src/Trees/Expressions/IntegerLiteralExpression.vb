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
''' A parse tree for an integer literal.
''' </summary>
Public NotInheritable Class IntegerLiteralExpression
    Inherits LiteralExpression

    Private ReadOnly _Literal As Integer
    Private ReadOnly _TypeCharacter As TypeCharacter
    Private ReadOnly _IntegerBase As IntegerBase

    ''' <summary>
    ''' The literal value.
    ''' </summary>
    Public ReadOnly Property Literal() As Integer
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
    ''' The type character on the literal.
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
    ''' Constructs a new parse tree for an integer literal.
    ''' </summary>
    ''' <param name="literal">The literal value.</param>
    ''' <param name="integerBase">The integer base of the literal.</param>
    ''' <param name="typeCharacter">The type character on the literal.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal literal As Integer, ByVal integerBase As IntegerBase, ByVal typeCharacter As TypeCharacter, ByVal span As Span)
        MyBase.New(TreeType.IntegerLiteralExpression, span)

        If integerBase < integerBase.Decimal OrElse integerBase > integerBase.Hexadecimal Then
            Throw New ArgumentOutOfRangeException("integerBase")
        End If

        If typeCharacter <> typeCharacter.None AndAlso _
           typeCharacter <> typeCharacter.IntegerSymbol AndAlso typeCharacter <> typeCharacter.IntegerChar AndAlso _
           typeCharacter <> typeCharacter.ShortChar AndAlso _
           typeCharacter <> typeCharacter.LongSymbol AndAlso typeCharacter <> typeCharacter.LongChar Then
            Throw New ArgumentOutOfRangeException("typeCharacter")
        End If

        _Literal = literal
        _IntegerBase = integerBase
        _TypeCharacter = typeCharacter
    End Sub
End Class