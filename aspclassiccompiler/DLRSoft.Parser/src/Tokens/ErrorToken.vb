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
''' A lexical error.
''' </summary>
Public NotInheritable Class ErrorToken
    Inherits Token

    Private ReadOnly _SyntaxError As SyntaxError

    ''' <summary>
    ''' The syntax error that represents the lexical error.
    ''' </summary>
    Public ReadOnly Property SyntaxError() As SyntaxError
        Get
            Return _SyntaxError
        End Get
    End Property

    ''' <summary>
    ''' Creates a new lexical error token.
    ''' </summary>
    ''' <param name="errorType">The type of the error.</param>
    ''' <param name="span">The location of the error.</param>
    Public Sub New(ByVal errorType As SyntaxErrorType, ByVal span As Span)
        MyBase.New(TokenType.LexicalError, span)

        If errorType < SyntaxErrorType.InvalidEscapedIdentifier OrElse errorType > SyntaxErrorType.InvalidDecimalLiteral Then
            Throw New ArgumentOutOfRangeException("errorType")
        End If

        _SyntaxError = New SyntaxError(errorType, span)
    End Sub
End Class