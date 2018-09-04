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
''' A punctuation token.
''' </summary>
Public NotInheritable Class PunctuatorToken
    Inherits Token

    Private Shared PunctuatorTable As Dictionary(Of String, TokenType)

    Private Shared Sub AddPunctuator(ByVal table As Dictionary(Of String, TokenType), ByVal punctuator As String, ByVal type As TokenType)
        table.Add(punctuator, type)
        table.Add(Scanner.MakeFullWidth(punctuator), type)
    End Sub

    ' Returns the token type of the string.
    Friend Shared Function TokenTypeFromString(ByVal s As String) As TokenType
        If PunctuatorTable Is Nothing Then
            Dim Table As New Dictionary(Of String, TokenType)(StringComparer.InvariantCulture)

            ' NOTE: These have to be in the same order as the enum!
            AddPunctuator(Table, "(", TokenType.LeftParenthesis)
            AddPunctuator(Table, ")", TokenType.RightParenthesis)
            AddPunctuator(Table, "{", TokenType.LeftCurlyBrace)
            AddPunctuator(Table, "}", TokenType.RightCurlyBrace)
            AddPunctuator(Table, "!", TokenType.Exclamation)
            AddPunctuator(Table, "#", TokenType.Pound)
            AddPunctuator(Table, ",", TokenType.Comma)
            AddPunctuator(Table, ".", TokenType.Period)
            AddPunctuator(Table, ":", TokenType.Colon)
            AddPunctuator(Table, ":=", TokenType.ColonEquals)
            AddPunctuator(Table, "&", TokenType.Ampersand)
            AddPunctuator(Table, "&=", TokenType.AmpersandEquals)
            AddPunctuator(Table, "*", TokenType.Star)
            AddPunctuator(Table, "*=", TokenType.StarEquals)
            AddPunctuator(Table, "+", TokenType.Plus)
            AddPunctuator(Table, "+=", TokenType.PlusEquals)
            AddPunctuator(Table, "-", TokenType.Minus)
            AddPunctuator(Table, "-=", TokenType.MinusEquals)
            AddPunctuator(Table, "/", TokenType.ForwardSlash)
            AddPunctuator(Table, "/=", TokenType.ForwardSlashEquals)
            AddPunctuator(Table, "\", TokenType.BackwardSlash)
            AddPunctuator(Table, "\=", TokenType.BackwardSlashEquals)
            AddPunctuator(Table, "^", TokenType.Caret)
            AddPunctuator(Table, "^=", TokenType.CaretEquals)
            AddPunctuator(Table, "<", TokenType.LessThan)
            AddPunctuator(Table, "<=", TokenType.LessThanEquals)
            AddPunctuator(Table, "=<", TokenType.LessThanEquals) 'lc VBScript allows the other way
            AddPunctuator(Table, "=", TokenType.Equals)
            AddPunctuator(Table, "<>", TokenType.NotEquals)
            AddPunctuator(Table, ">", TokenType.GreaterThan)
            AddPunctuator(Table, ">=", TokenType.GreaterThanEquals)
            AddPunctuator(Table, "=>", TokenType.GreaterThanEquals) 'lc
            AddPunctuator(Table, "<<", TokenType.LessThanLessThan)
            AddPunctuator(Table, "<<=", TokenType.LessThanLessThanEquals)
            AddPunctuator(Table, ">>", TokenType.GreaterThanGreaterThan)
            AddPunctuator(Table, ">>=", TokenType.GreaterThanGreaterThanEquals)

            PunctuatorTable = Table
        End If

        If Not PunctuatorTable.ContainsKey(s) Then
            Return TokenType.None
        Else
            Return PunctuatorTable(s)
        End If
    End Function

    ''' <summary>
    ''' Constructs a new punctuator token.
    ''' </summary>
    ''' <param name="type">The punctuator token type.</param>
    ''' <param name="span">The location of the punctuator.</param>
    Public Sub New(ByVal type As TokenType, ByVal span As Span)
        MyBase.New(type, span)

        If (type < TokenType.LeftParenthesis OrElse type > TokenType.GreaterThanGreaterThanEquals) Then
            Throw New ArgumentOutOfRangeException("type")
        End If
    End Sub
End Class