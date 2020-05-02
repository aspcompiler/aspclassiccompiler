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
''' A parse tree for a simple name (e.g. 'foo').
''' </summary>
Public NotInheritable Class SimpleName
    Inherits Name

    Private ReadOnly _Name As String
    Private ReadOnly _TypeCharacter As TypeCharacter
    Private ReadOnly _Escaped As Boolean

    ''' <summary>
    ''' The name, if any.
    ''' </summary>
    Public ReadOnly Property Name() As String
        Get
            Return _Name
        End Get
    End Property

    ''' <summary>
    ''' The type character.
    ''' </summary>
    Public ReadOnly Property TypeCharacter() As TypeCharacter
        Get
            Return _TypeCharacter
        End Get
    End Property

    ''' <summary>
    ''' Whether the name is escaped.
    ''' </summary>
    Public ReadOnly Property Escaped() As Boolean
        Get
            Return _Escaped
        End Get
    End Property

    ''' <summary>
    ''' Creates a bad simple name.
    ''' </summary>
    ''' <param name="Span">The location of the parse tree.</param>
    ''' <returns>A bad simple name.</returns>
    Public Shared Function GetBadSimpleName(ByVal span As Span) As SimpleName
        Return New SimpleName(Span)
    End Function

    ''' <summary>
    ''' Constructs a new simple name parse tree.
    ''' </summary>
    ''' <param name="name">The name, if any.</param>
    ''' <param name="typeCharacter">The type character.</param>
    ''' <param name="escaped">Whether the name is escaped.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal name As String, ByVal typeCharacter As TypeCharacter, ByVal escaped As Boolean, ByVal span As Span)
        MyBase.New(TreeType.SimpleName, span)

        If typeCharacter <> typeCharacter.None AndAlso escaped Then
            Throw New ArgumentException("Escaped named cannot have type characters.")
        End If

        If typeCharacter <> typeCharacter.None AndAlso typeCharacter <> typeCharacter.DecimalSymbol AndAlso _
           typeCharacter <> typeCharacter.DoubleSymbol AndAlso typeCharacter <> typeCharacter.IntegerSymbol AndAlso _
           typeCharacter <> typeCharacter.LongSymbol AndAlso typeCharacter <> typeCharacter.SingleSymbol AndAlso _
           typeCharacter <> typeCharacter.StringSymbol Then
            Throw New ArgumentOutOfRangeException("typeCharacter")
        End If

        If name Is Nothing Then
            Throw New ArgumentNullException("name")
        End If

        _Name = name
        _TypeCharacter = typeCharacter
        _Escaped = escaped
    End Sub

    Private Sub New(ByVal span As Span)
        MyBase.New(TreeType.SimpleName, span)
    End Sub

    Public Overrides ReadOnly Property IsBad() As Boolean
        Get
            Return Name Is Nothing
        End Get
    End Property
End Class