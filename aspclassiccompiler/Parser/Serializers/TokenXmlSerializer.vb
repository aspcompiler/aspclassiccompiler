'
' Visual Basic .NET Parser
'
' Copyright (C) 2005, Microsoft Corporation. All rights reserved.
'
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
' MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
'

Imports System.Xml

Public Class TokenXmlSerializer
    Private ReadOnly Writer As XmlWriter

    Public Sub New(ByVal Writer As XmlWriter)
        Me.Writer = Writer
    End Sub

    Private Sub Serialize(ByVal Span As Span)
        Writer.WriteAttributeString("startLine", CStr(Span.Start.Line))
        Writer.WriteAttributeString("startCol", CStr(Span.Start.Column))
        Writer.WriteAttributeString("endLine", CStr(Span.Finish.Line))
        Writer.WriteAttributeString("endCol", CStr(Span.Finish.Column))
    End Sub

    Private Sub Serialize(ByVal TypeCharacter As TypeCharacter)
        If TypeCharacter <> TypeCharacter.None Then
            Static TypeCharacterTable As Dictionary(Of TypeCharacter, String)

            If TypeCharacterTable Is Nothing Then
                Dim Table As Dictionary(Of TypeCharacter, String) = New Dictionary(Of TypeCharacter, String)()
                ' NOTE: These have to be in the same order as the enum!
                Dim TypeCharacters() As String = {"$", "%", "&", "S", "I", "L", "!", "#", "@", "F", "R", "D", "US", "UI", "UL"}
                Dim TableTypeCharacter As TypeCharacter = TypeCharacter.StringSymbol

                For Index As Integer = 0 To TypeCharacters.Length - 1
                    Table.Add(TableTypeCharacter, TypeCharacters(Index))
                    TableTypeCharacter = CType(TableTypeCharacter << 1, TypeCharacter)
                Next

                TypeCharacterTable = Table
            End If

            Writer.WriteAttributeString("typeChar", TypeCharacterTable(TypeCharacter))
        End If
    End Sub

    Public Sub Serialize(ByVal Token As Token)
        Writer.WriteStartElement(Token.Type.ToString())
        Serialize(Token.Span)

        Select Case Token.Type
            Case TokenType.LexicalError
                With CType(Token, ErrorToken)
                    Writer.WriteAttributeString("errorNumber", CStr(.SyntaxError.Type))
                    Writer.WriteString(.SyntaxError.ToString())
                End With

            Case TokenType.Comment
                With CType(Token, CommentToken)
                    Writer.WriteAttributeString("isRem", CStr(.IsREM))
                    Writer.WriteString(.Comment)
                End With

            Case TokenType.Identifier
                With CType(Token, IdentifierToken)
                    Writer.WriteAttributeString("escaped", CStr(.Escaped))
                    Serialize(.TypeCharacter)
                    Writer.WriteString(.Identifier)
                End With

            Case TokenType.StringLiteral
                With CType(Token, StringLiteralToken)
                    Writer.WriteString(.Literal)
                End With

            Case TokenType.CharacterLiteral
                With CType(Token, CharacterLiteralToken)
                    Writer.WriteString(.Literal)
                End With

            Case TokenType.DateLiteral
                With CType(Token, DateLiteralToken)
                    Writer.WriteString(CStr(.Literal))
                End With

            Case TokenType.IntegerLiteral
                With CType(Token, IntegerLiteralToken)
                    Writer.WriteAttributeString("base", .IntegerBase.ToString())
                    Serialize(.TypeCharacter)
                    Writer.WriteString(CStr(.Literal))
                End With

            Case TokenType.FloatingPointLiteral
                With CType(Token, FloatingPointLiteralToken)
                    Serialize(.TypeCharacter)
                    Writer.WriteString(CStr(.Literal))
                End With

            Case TokenType.DecimalLiteral
                With CType(Token, DecimalLiteralToken)
                    Serialize(.TypeCharacter)
                    Writer.WriteString(CStr(.Literal))
                End With

            Case TokenType.UnsignedIntegerLiteral
                With CType(Token, UnsignedIntegerLiteralToken)
                    Serialize(.TypeCharacter)
                    Writer.WriteString(CStr(.Literal))
                End With

            Case Else
                ' Fall through
        End Select

        Writer.WriteEndElement()
    End Sub

    Public Sub Serialize(ByVal Tokens() As Token)
        For Each Token As Token In Tokens
            Serialize(Token)
        Next
    End Sub

End Class
