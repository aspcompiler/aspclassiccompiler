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
''' A string literal.
''' </summary>
Public NotInheritable Class StringLiteralToken
    Inherits Token

    Private ReadOnly _Literal As String

    ''' <summary>
    ''' The value of the literal.
    ''' </summary>
    Public ReadOnly Property Literal() As String
        Get
            Return _Literal
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new string literal token.
    ''' </summary>
    ''' <param name="literal">The value of the literal.</param>
    ''' <param name="span">The location of the literal.</param>
    Public Sub New(ByVal literal As String, ByVal span As Span)
        MyBase.New(TokenType.StringLiteral, span)

        If literal Is Nothing Then
            Throw New ArgumentNullException("literal")
        End If

        _Literal = literal
    End Sub
End Class