'
' Visual Basic .NET Parser
'
' Copyright (C) 2005, Microsoft Corporation. All rights reserved.
'
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
' MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
'

Imports System.Globalization

''' <summary>
''' A lexical analyzer for Visual Basic .NET. It produces a stream of lexical tokens.
''' </summary>
Public NotInheritable Class Scanner
    Implements IDisposable

    ' The text to be read. We use a TextReader here so that lexical analysis
    ' can be done on strings as well as streams.
    Private _Source As TextReader

    ' For performance reasons, we cache the character when we peek ahead.
    Private _PeekCache As Char
    Private _PeekCacheHasValue As Boolean = False

    ' There are a few places where we're going to need to peek one character
    ' ahead
    Private _PeekAheadCache As Char
    Private _PeekAheadCacheHasValue As Boolean = False

    ' Since we're only using a TextReader which has no position information,
    ' we have to keep track of line/column information ourselves.
    Private _Index As Integer = 0
    Private _Line As Integer = 1
    Private _Column As Integer = 1

    ' A buffer of all the tokens we've returned so far
    Private _Tokens As List(Of Token) = New List(Of Token)()

    ' Our current position in the buffer. -1 means before the beginning.
    Private _Position As Integer = -1

    ' Determine whether we have been disposed already or not
    Private _Disposed As Boolean = False

    ' How many columns a tab character should be treated as
    Private _TabSpaces As Integer = 4

    ' Version of the language to parse
    Private _Version As LanguageVersion = LanguageVersion.VisualBasic80

    ''' <summary>
    ''' How many columns a tab character should be considered.
    ''' </summary>
    Public Property TabSpaces() As Integer
        Get
            Return _TabSpaces
        End Get

        Set(ByVal value As Integer)
            If Value < 1 Then
                Throw New ArgumentException("Tabs cannot represent less than one space.")
            End If

            _TabSpaces = Value
        End Set
    End Property

    ''' <summary>
    ''' The version of Visual Basic this scanner operates on.
    ''' </summary>
    Public ReadOnly Property Version() As LanguageVersion
        Get
            Return _Version
        End Get
    End Property

    ''' <summary>
    ''' Constructs a scanner for a string.
    ''' </summary>
    ''' <param name="source">The string to scan.</param>
    Public Sub New(ByVal source As String)
        If source Is Nothing Then Throw New ArgumentNullException("Source")
        _Source = New StringReader(source)
    End Sub

    ''' <summary>
    ''' Constructs a scanner for a string.
    ''' </summary>
    ''' <param name="source">The string to scan.</param>
    ''' <param name="version">The language version to parse.</param>
    Public Sub New(ByVal source As String, ByVal version As LanguageVersion)
        If source Is Nothing Then Throw New ArgumentNullException("Source")
        If version <> LanguageVersion.VisualBasic71 AndAlso version <> LanguageVersion.VisualBasic80 Then
            Throw New ArgumentOutOfRangeException("Version")
        End If
        _Source = New StringReader(source)
        _Version = version
    End Sub

    ''' <summary>
    ''' Constructs a scanner for a stream.
    ''' </summary>
    ''' <param name="source">The stream to scan.</param>
    Public Sub New(ByVal source As Stream)
        If source Is Nothing Then Throw New ArgumentNullException("Source")
        _Source = New StreamReader(source)
    End Sub

    ''' <summary>
    ''' Constructs a scanner for a stream.
    ''' </summary>
    ''' <param name="source">The stream to scan.</param>
    ''' <param name="version">The language version to parse.</param>
    Public Sub New(ByVal source As Stream, ByVal version As LanguageVersion)
        If source Is Nothing Then Throw New ArgumentNullException("Source")
        If version <> LanguageVersion.VisualBasic71 AndAlso version <> LanguageVersion.VisualBasic80 Then
            Throw New ArgumentOutOfRangeException("Version")
        End If
        _Source = New StreamReader(source)
        _Version = version
    End Sub

    ''' <summary>
    ''' Constructs a canner for a general TextReader.
    ''' </summary>
    ''' <param name="source">The TextReader to scan.</param>
    Public Sub New(ByVal source As TextReader)
        If source Is Nothing Then Throw New ArgumentNullException("Source")
        _Source = source
    End Sub

    ''' <summary>
    ''' Constructs a canner for a general TextReader.
    ''' </summary>
    ''' <param name="source">The TextReader to scan.</param>
    ''' <param name="version">The language version to parse.</param>
    Public Sub New(ByVal source As TextReader, ByVal version As LanguageVersion)
        If source Is Nothing Then Throw New ArgumentNullException("Source")
        If version <> LanguageVersion.VisualBasic71 AndAlso version <> LanguageVersion.VisualBasic80 Then
            Throw New ArgumentOutOfRangeException("Version")
        End If
        _Source = source
        _Version = version
    End Sub

    ''' <summary>
    ''' Closes/disposes the scanner.
    ''' </summary>
    Public Sub Close() Implements IDisposable.Dispose
        If Not _Disposed Then
            _Disposed = True
            _Source.Close()
        End If
    End Sub

    ' Read a character
    Private Function ReadChar() As Char
        Dim c As Char

        If _PeekCacheHasValue Then
            c = _PeekCache
            _PeekCacheHasValue = False

            If _PeekAheadCacheHasValue Then
                _PeekCache = _PeekAheadCache
                _PeekCacheHasValue = True
                _PeekAheadCacheHasValue = False
            End If
        Else
            Debug.Assert(Not _PeekAheadCacheHasValue, "Cache incorrect!")
            c = ChrW(_Source.Read())
        End If

        _Index += 1
        If AscW(c) = &H9 Then
            _Column += _TabSpaces
        Else
            _Column += 1
        End If

        Return c
    End Function

    ' Peek ahead at the next character
    Private Function PeekChar() As Char
        If Not _PeekCacheHasValue Then
            _PeekCache = ChrW(_Source.Read())
            _PeekCacheHasValue = True
        End If

        Return _PeekCache
    End Function

    ' Peek at the character past the next character
    Private Function PeekAheadChar() As Char
        If Not _PeekAheadCacheHasValue Then
            If Not _PeekCacheHasValue Then
                PeekChar()
            End If

            _PeekAheadCache = ChrW(_Source.Read())
            _PeekAheadCacheHasValue = True
        End If

        Return _PeekAheadCache
    End Function

    ' The current line/column position
    Private ReadOnly Property CurrentLocation() As Location
        Get
            Return New Location(_Index, _Line, _Column)
        End Get
    End Property

    ' Creates a span from the start location to the current location.
    Private Function SpanFrom(ByVal start As Location) As Span
        Return New Span(start, CurrentLocation)
    End Function

    Private Shared Function IsAlphaClass(ByVal c As UnicodeCategory) As Boolean
        Return c = UnicodeCategory.UppercaseLetter OrElse _
               c = UnicodeCategory.LowercaseLetter OrElse _
               c = UnicodeCategory.TitlecaseLetter OrElse _
               c = UnicodeCategory.OtherLetter OrElse _
               c = UnicodeCategory.ModifierLetter OrElse _
               c = UnicodeCategory.LetterNumber
    End Function

    Private Shared Function IsNumericClass(ByVal c As UnicodeCategory) As Boolean
        Return c = UnicodeCategory.DecimalDigitNumber
    End Function

    Private Shared Function IsUnderscoreClass(ByVal c As UnicodeCategory) As Boolean
        Return c = UnicodeCategory.ConnectorPunctuation
    End Function

    Private Shared Function IsSingleQuote(ByVal c As Char) As Boolean
        Return c = "'"c OrElse c = ChrW(&HFF07) OrElse c = ChrW(&H2018) OrElse c = ChrW(&H2019)
    End Function

    Private Shared Function IsDoubleQuote(ByVal c As Char) As Boolean
        Return c = """"c OrElse c = ChrW(&HFF02) OrElse c = ChrW(&H201C) OrElse c = ChrW(&H201D)
    End Function

    Private Shared Function IsDigit(ByVal c As Char) As Boolean
        Return (c >= "0"c AndAlso c <= "9"c) OrElse (c >= ChrW(&HFF10) AndAlso c <= ChrW(&HFF19))
    End Function

    Private Shared Function IsOctalDigit(ByVal c As Char) As Boolean
        Return (c >= "0"c AndAlso c <= "7"c) OrElse (c >= ChrW(&HFF10) AndAlso c <= ChrW(&HFF17))
    End Function

    Private Shared Function IsHexDigit(ByVal c As Char) As Boolean
        Return IsDigit(c) OrElse _
               (c >= "a"c AndAlso c <= "f"c) OrElse (c >= "A"c AndAlso c <= "F"c) OrElse _
               (c >= ChrW(&HFF41) AndAlso c <= ChrW(&HFF46)) OrElse (c >= ChrW(&HFF21) AndAlso c <= ChrW(&HFF26))
    End Function

    Private Shared Function IsEquals(ByVal c As Char) As Boolean
        Return c = "="c OrElse c = ChrW(&HFF1D)
    End Function

    Private Shared Function IsLessThan(ByVal c As Char) As Boolean
        Return c = "<"c OrElse c = ChrW(&HFF1C)
    End Function

    Private Shared Function IsGreaterThan(ByVal c As Char) As Boolean
        Return c = ">"c OrElse c = ChrW(&HFF1E)
    End Function

    Private Shared Function IsAmpersand(ByVal c As Char) As Boolean
        Return c = "&"c OrElse c = ChrW(&HFF06)
    End Function

    Private Shared Function IsUnderscore(ByVal c As Char) As Boolean
        Return IsUnderscoreClass(Char.GetUnicodeCategory(c))
    End Function

    Private Shared Function IsHexDesignator(ByVal c As Char) As Boolean
        Return c = "H"c OrElse c = "h"c OrElse c = ChrW(&HFF48) OrElse c = ChrW(&HFF28)
    End Function

    Private Shared Function IsOctalDesignator(ByVal c As Char) As Boolean
        Return c = "O"c OrElse c = "o"c OrElse c = ChrW(&HFF2F) OrElse c = ChrW(&HFF4F)
    End Function

    Private Shared Function IsPeriod(ByVal c As Char) As Boolean
        Return c = "."c OrElse c = ChrW(&HFF0E)
    End Function

    Private Shared Function IsExponentDesignator(ByVal c As Char) As Boolean
        Return c = "e"c OrElse c = "E"c OrElse c = ChrW(&HFF45) OrElse c = ChrW(&HFF25)
    End Function

    Private Shared Function IsPlus(ByVal c As Char) As Boolean
        Return c = "+"c OrElse c = ChrW(&HFF0B)
    End Function

    Private Shared Function IsMinus(ByVal c As Char) As Boolean
        Return c = "-"c OrElse c = ChrW(&HFF0D)
    End Function

    Private Shared Function IsForwardSlash(ByVal c As Char) As Boolean
        Return c = "/"c OrElse c = ChrW(&HFF0F)
    End Function

    Private Shared Function IsColon(ByVal c As Char) As Boolean
        Return c = ":"c OrElse c = ChrW(&HFF1A)
    End Function

    Private Shared Function IsPound(ByVal c As Char) As Boolean
        Return c = "#"c OrElse c = ChrW(&HFF03)
    End Function

    Private Shared Function IsA(ByVal c As Char) As Boolean
        Return c = "a"c OrElse c = ChrW(&HFF41) OrElse c = "A"c OrElse c = ChrW(&HFF21)
    End Function

    Private Shared Function IsP(ByVal c As Char) As Boolean
        Return c = "p"c OrElse c = ChrW(&HFF50) OrElse c = "P"c OrElse c = ChrW(&HFF30)
    End Function

    Private Shared Function IsM(ByVal c As Char) As Boolean
        Return c = "m"c OrElse c = ChrW(&HFF4D) OrElse c = "M"c OrElse c = ChrW(&HFF2D)
    End Function

    Private Shared Function IsCharDesignator(ByVal c As Char) As Boolean
        Return c = "c"c OrElse c = "C"c OrElse c = ChrW(&HFF43) OrElse c = ChrW(&HFF23)
    End Function

    Private Shared Function IsLeftBracket(ByVal c As Char) As Boolean
        Return c = "["c OrElse c = ChrW(&HFF3B)
    End Function

    Private Shared Function IsRightBracket(ByVal c As Char) As Boolean
        Return c = "]"c OrElse c = ChrW(&HFF3D)
    End Function

    Private Shared Function IsUnsignedTypeChar(ByVal c As Char) As Boolean
        Return c = "u"c OrElse c = "U"c OrElse c = ChrW(&HFF35) OrElse c = ChrW(&HFF55)
    End Function

    Private Shared Function IsIdentifier(ByVal c As Char) As Boolean
        Dim CharClass As UnicodeCategory = Char.GetUnicodeCategory(c)

        Return _
            IsAlphaClass(CharClass) OrElse _
            IsNumericClass(CharClass) OrElse _
            CharClass = UnicodeCategory.SpacingCombiningMark OrElse _
            CharClass = UnicodeCategory.NonSpacingMark OrElse _
            CharClass = UnicodeCategory.Format OrElse _
            IsUnderscoreClass(CharClass)
    End Function

    Friend Shared Function MakeHalfWidth(ByVal c As Char) As Char
        If c < ChrW(&HFF01) OrElse c > ChrW(&HFF5E) Then
            Return c
        Else
            Return ChrW(AscW(c) - &HFF00 + &H20)
        End If
    End Function

    Friend Shared Function MakeFullWidth(ByVal c As Char) As Char
        If c < ChrW(&H21) OrElse c > ChrW(&H7E) Then
            Return c
        Else
            Return ChrW(AscW(c) + &HFF00 - &H20)
        End If
    End Function

    Friend Shared Function MakeFullWidth(ByVal s As String) As String
        Dim Builder As StringBuilder = New StringBuilder(s)

        For Index As Integer = 0 To Builder.Length - 1
            Builder(Index) = MakeFullWidth(Builder(Index))
        Next

        Return Builder.ToString()
    End Function

    '
    ' Scan functions
    '
    ' Each function assumes that the reader is positioned at the beginning of
    ' the token. At the end, the function will have read through the entire
    ' token. If an error occurs, the function may attempt to do error recovery.
    '

    Private Function ScanPossibleTypeCharacter(ByVal ValidTypeCharacters As TypeCharacter) As TypeCharacter
        Dim TypeChar As Char = PeekChar()
        Dim TypeString As String
        Static TypeCharacterTable As Dictionary(Of String, TypeCharacter)

        If TypeCharacterTable Is Nothing Then
            Dim Table As Dictionary(Of String, TypeCharacter) = New Dictionary(Of String, TypeCharacter)(StringComparer.InvariantCultureIgnoreCase)
            ' NOTE: These have to be in the same order as the enum!
            Dim TypeCharacters() As String = {"$", "%", "&", "S", "I", "L", "!", "#", "@", "F", "R", "D", "US", "UI", "UL"}
            Dim TypeCharacter As TypeCharacter = TypeCharacter.StringSymbol

            For Index As Integer = 0 To TypeCharacters.Length - 1
                With Table
                    .Add(TypeCharacters(Index), TypeCharacter)
                    .Add(Scanner.MakeFullWidth(TypeCharacters(Index)), TypeCharacter)
                End With

                TypeCharacter = CType(TypeCharacter << 1, TypeCharacter)
            Next

            TypeCharacterTable = Table
        End If

        If IsUnsignedTypeChar(TypeChar) AndAlso _Version > LanguageVersion.VisualBasic71 Then
            ' At the point at which we've seen a "U", we don't know if it's going to
            ' be a valid type character or just something invalid.
            TypeString = TypeChar & PeekAheadChar()
        Else
            TypeString = TypeChar
        End If

        If TypeCharacterTable.ContainsKey(TypeString) Then
            Dim TypeCharacter As TypeCharacter = TypeCharacterTable(TypeString)

            If (TypeCharacter And ValidTypeCharacters) <> 0 Then
                ' A bang (!) is a type character unless it is followed by a legal identifier start.
                If TypeCharacter = TypeCharacter.SingleSymbol AndAlso CanStartIdentifier(PeekAheadChar()) Then
                    Return TypeCharacter.None
                End If

                ReadChar()

                If IsUnsignedTypeChar(TypeChar) Then
                    ReadChar()
                End If

                Return TypeCharacter
            End If
        End If

        Return TypeCharacter.None
    End Function

    Private Function ScanPossibleMultiCharacterPunctuator(ByVal leadingCharacter As Char, ByVal start As Location) As PunctuatorToken
        Dim NextChar As Char = PeekChar()
        Dim Punctuator As TokenType
        Dim PunctuatorString As String = leadingCharacter

        Debug.Assert(PunctuatorToken.TokenTypeFromString(leadingCharacter) <> TokenType.None)

        If IsEquals(NextChar) OrElse IsLessThan(NextChar) OrElse IsGreaterThan(NextChar) Then
            PunctuatorString &= NextChar
            Punctuator = PunctuatorToken.TokenTypeFromString(PunctuatorString)

            If Punctuator <> TokenType.None Then
                ReadChar()

                If (Punctuator = TokenType.LessThanLessThan OrElse _
                    Punctuator = TokenType.GreaterThanGreaterThan) AndAlso _
                   IsEquals(PeekChar()) Then
                    PunctuatorString &= ReadChar()
                    Punctuator = PunctuatorToken.TokenTypeFromString(PunctuatorString)
                End If

                Return New PunctuatorToken(Punctuator, SpanFrom(start))
            End If
        End If

        Punctuator = PunctuatorToken.TokenTypeFromString(leadingCharacter)
        Return New PunctuatorToken(Punctuator, SpanFrom(start))
    End Function

    Private Function ScanNumericLiteral() As Token
        Dim Start As Location = CurrentLocation
        Dim Literal As StringBuilder = New StringBuilder()
        Dim Base As IntegerBase = IntegerBase.Decimal
        Dim TypeCharacter As TypeCharacter = TypeCharacter.None

        Debug.Assert(CanStartNumericLiteral())

        If IsAmpersand(PeekChar()) Then
            Literal.Append(MakeHalfWidth(ReadChar()))

            If IsHexDesignator(PeekChar()) Then
                Literal.Append(MakeHalfWidth(ReadChar()))
                Base = IntegerBase.Hexadecimal

                While IsHexDigit(PeekChar())
                    Literal.Append(MakeHalfWidth(ReadChar()))
                End While
            ElseIf IsOctalDesignator(PeekChar()) Then
                Literal.Append(MakeHalfWidth(ReadChar()))
                Base = IntegerBase.Octal

                While IsOctalDigit(PeekChar())
                    Literal.Append(MakeHalfWidth(ReadChar()))
                End While
            ElseIf IsOctalDigit(PeekChar()) Then 'VbScript Octal is like &123456&
                Base = IntegerBase.Octal
                Literal.Append("O"c) 'VB.net Octal starts with &O

                While IsOctalDigit(PeekChar())
                    Literal.Append(MakeHalfWidth(ReadChar()))
                End While

                If IsAmpersand(PeekChar()) Then
                    ReadChar() 'Ignored the last '&'
                End If
            Else
                Return ScanPossibleMultiCharacterPunctuator("&"c, Start)
            End If

            If Literal.Length > 2 Then
                Const ValidTypeChars As TypeCharacter = _
                    TypeCharacter.ShortChar Or TypeCharacter.UnsignedShortChar Or _
                    TypeCharacter.IntegerSymbol Or TypeCharacter.IntegerChar Or TypeCharacter.UnsignedIntegerChar Or _
                    TypeCharacter.LongSymbol Or TypeCharacter.LongChar Or TypeCharacter.UnsignedLongChar

                TypeCharacter = ScanPossibleTypeCharacter(ValidTypeChars)

                Try
                    Select Case TypeCharacter
                        Case TypeCharacter.ShortChar
                            Dim Value As Long = CLng(Literal.ToString())

                            If Value <= &HFFFFL Then
                                If Value > &H7FFFL Then
                                    Value = -(&H10000L - Value)
                                End If

                                If Value >= Short.MinValue AndAlso Value <= Short.MaxValue Then
                                    Return New IntegerLiteralToken(CShort(Value), Base, TypeCharacter, SpanFrom(Start))
                                End If
                            End If
                            ' Fall through

                        Case TypeCharacter.UnsignedShortChar
                            Dim Value As ULong = CULng(Literal.ToString())

                            If Value <= &HFFFFL Then
                                If Value >= UShort.MinValue AndAlso Value <= UShort.MaxValue Then
                                    Return New UnsignedIntegerLiteralToken(CUShort(Value), Base, TypeCharacter, SpanFrom(Start))
                                End If
                            End If
                            ' Fall through

                        Case TypeCharacter.IntegerSymbol, TypeCharacter.IntegerChar
                            Dim Value As Long = CLng(Literal.ToString())

                            If Value <= &HFFFFFFFFL Then
                                If Value > &H7FFFFFFFL Then
                                    Value = -(&H100000000L - Value)
                                End If

                                If Value >= Integer.MinValue AndAlso Value <= Integer.MaxValue Then
                                    Return New IntegerLiteralToken(CInt(Value), Base, TypeCharacter, SpanFrom(Start))
                                End If
                            End If
                            ' Fall through

                        Case TypeCharacter.UnsignedIntegerChar
                            Dim Value As ULong = CULng(Literal.ToString())

                            If Value <= &HFFFFFFFFL Then
                                If Value >= UInteger.MinValue AndAlso Value <= UInteger.MaxValue Then
                                    Return New UnsignedIntegerLiteralToken(CUInt(Value), Base, TypeCharacter, SpanFrom(Start))
                                End If
                            End If
                            ' Fall through

                        Case TypeCharacter.LongSymbol, TypeCharacter.LongChar
                            Return New IntegerLiteralToken(ParseInt(Literal.ToString()), Base, TypeCharacter, SpanFrom(Start))

                        Case TypeCharacter.UnsignedLongChar
                            Return New UnsignedIntegerLiteralToken(CULng(Literal.ToString()), Base, TypeCharacter, SpanFrom(Start))

                        Case Else
                            TypeCharacter = TypeCharacter.None
                            Return New IntegerLiteralToken(ParseInt(Literal.ToString()), Base, TypeCharacter, SpanFrom(Start))
                    End Select
                Catch ex As OverflowException
                    Return New ErrorToken(SyntaxErrorType.InvalidIntegerLiteral, SpanFrom(Start))
                Catch ex As InvalidCastException
                    Return New ErrorToken(SyntaxErrorType.InvalidIntegerLiteral, SpanFrom(Start))
                End Try
            End If

            Return New ErrorToken(SyntaxErrorType.InvalidIntegerLiteral, SpanFrom(Start))
        End If

        While IsDigit(PeekChar())
            Literal.Append(MakeHalfWidth(ReadChar()))
        End While

        If IsPeriod(PeekChar()) OrElse IsExponentDesignator(PeekChar()) Then
            Dim ErrorType As SyntaxErrorType = SyntaxErrorType.None
            Const ValidTypeChars As TypeCharacter = _
                TypeCharacter.DecimalChar Or TypeCharacter.DecimalSymbol Or _
                TypeCharacter.SingleChar Or TypeCharacter.SingleSymbol Or _
                TypeCharacter.DoubleChar Or TypeCharacter.DoubleSymbol

            If IsPeriod(PeekChar()) Then
                Literal.Append(MakeHalfWidth(ReadChar()))

                If Not IsDigit(PeekChar()) And Literal.Length = 1 Then
                    Return New PunctuatorToken(TokenType.Period, SpanFrom(Start))
                End If

                While IsDigit(PeekChar())
                    Literal.Append(MakeHalfWidth(ReadChar()))
                End While
            End If

            If IsExponentDesignator(PeekChar()) Then
                Literal.Append(MakeHalfWidth(ReadChar()))

                If IsPlus(PeekChar()) OrElse IsMinus(PeekChar()) Then
                    Literal.Append(MakeHalfWidth(ReadChar()))
                End If

                If Not IsDigit(PeekChar()) Then
                    Return New ErrorToken(SyntaxErrorType.InvalidFloatingPointLiteral, SpanFrom(Start))
                End If

                While IsDigit(PeekChar())
                    Literal.Append(MakeHalfWidth(ReadChar()))
                End While
            End If

            TypeCharacter = ScanPossibleTypeCharacter(ValidTypeChars)

            Try
                Select Case TypeCharacter
                    Case TypeCharacter.DecimalChar, TypeCharacter.DecimalSymbol
                        ErrorType = SyntaxErrorType.InvalidDecimalLiteral
                        Return New DecimalLiteralToken(CDec(Literal.ToString()), TypeCharacter, SpanFrom(Start))

                    Case TypeCharacter.SingleSymbol, TypeCharacter.SingleChar
                        ErrorType = SyntaxErrorType.InvalidFloatingPointLiteral
                        Return New FloatingPointLiteralToken(CSng(Literal.ToString()), TypeCharacter, SpanFrom(Start))

                    Case TypeCharacter.DoubleSymbol, TypeCharacter.DoubleChar
                        ErrorType = SyntaxErrorType.InvalidFloatingPointLiteral
                        Return New FloatingPointLiteralToken(CDbl(Literal.ToString()), TypeCharacter, SpanFrom(Start))

                    Case Else
                        ErrorType = SyntaxErrorType.InvalidFloatingPointLiteral
                        TypeCharacter = TypeCharacter.None
                        Return New FloatingPointLiteralToken(CDbl(Literal.ToString()), TypeCharacter, SpanFrom(Start))
                End Select
            Catch ex As OverflowException
                Return New ErrorToken(ErrorType, SpanFrom(Start))
            Catch ex As InvalidCastException
                Return New ErrorToken(ErrorType, SpanFrom(Start))
            End Try
        Else
            Dim ErrorType As SyntaxErrorType = SyntaxErrorType.None
            Const ValidTypeChars As TypeCharacter = _
                TypeCharacter.ShortChar Or _
                TypeCharacter.IntegerSymbol Or TypeCharacter.IntegerChar Or _
                TypeCharacter.LongSymbol Or TypeCharacter.LongChar Or _
                TypeCharacter.DecimalSymbol Or TypeCharacter.DecimalChar Or _
                TypeCharacter.SingleSymbol Or TypeCharacter.SingleChar Or _
                TypeCharacter.DoubleSymbol Or TypeCharacter.DoubleChar Or _
                TypeCharacter.UnsignedShortChar Or TypeCharacter.UnsignedIntegerChar Or _
                TypeCharacter.UnsignedLongChar

            TypeCharacter = ScanPossibleTypeCharacter(ValidTypeChars)

            Try
                Select Case TypeCharacter
                    Case TypeCharacter.ShortChar
                        ErrorType = SyntaxErrorType.InvalidIntegerLiteral
                        Return New IntegerLiteralToken(CShort(Literal.ToString()), Base, TypeCharacter, SpanFrom(Start))

                    Case TypeCharacter.UnsignedShortChar
                        ErrorType = SyntaxErrorType.InvalidIntegerLiteral
                        Return New UnsignedIntegerLiteralToken(CUShort(Literal.ToString()), Base, TypeCharacter, SpanFrom(Start))

                    Case TypeCharacter.IntegerSymbol, TypeCharacter.IntegerChar
                        ErrorType = SyntaxErrorType.InvalidIntegerLiteral
                        Return New IntegerLiteralToken(CInt(Literal.ToString()), Base, TypeCharacter, SpanFrom(Start))

                    Case TypeCharacter.UnsignedIntegerChar
                        ErrorType = SyntaxErrorType.InvalidIntegerLiteral
                        Return New UnsignedIntegerLiteralToken(CUInt(Literal.ToString()), Base, TypeCharacter, SpanFrom(Start))

                    Case TypeCharacter.LongSymbol, TypeCharacter.LongChar
                        ErrorType = SyntaxErrorType.InvalidIntegerLiteral
                        Return New IntegerLiteralToken(CInt(Literal.ToString()), Base, TypeCharacter, SpanFrom(Start))

                    Case TypeCharacter.UnsignedLongChar
                        ErrorType = SyntaxErrorType.InvalidIntegerLiteral
                        Return New UnsignedIntegerLiteralToken(CULng(Literal.ToString()), Base, TypeCharacter, SpanFrom(Start))

                    Case TypeCharacter.DecimalChar, TypeCharacter.DecimalSymbol
                        ErrorType = SyntaxErrorType.InvalidDecimalLiteral
                        Return New DecimalLiteralToken(CDec(Literal.ToString()), TypeCharacter, SpanFrom(Start))

                    Case TypeCharacter.SingleSymbol, TypeCharacter.SingleChar
                        ErrorType = SyntaxErrorType.InvalidFloatingPointLiteral
                        Return New FloatingPointLiteralToken(CSng(Literal.ToString()), TypeCharacter, SpanFrom(Start))

                    Case TypeCharacter.DoubleSymbol, TypeCharacter.DoubleChar
                        ErrorType = SyntaxErrorType.InvalidFloatingPointLiteral
                        Return New FloatingPointLiteralToken(CDbl(Literal.ToString()), TypeCharacter, SpanFrom(Start))

                    Case Else
                        ErrorType = SyntaxErrorType.InvalidIntegerLiteral
                        Return New IntegerLiteralToken(CInt(Literal.ToString()), Base, TypeCharacter.None, SpanFrom(Start))
                End Select
            Catch ex As OverflowException
                Return New ErrorToken(ErrorType, SpanFrom(Start))
            Catch ex As InvalidCastException
                Return New ErrorToken(ErrorType, SpanFrom(Start))
            End Try
        End If
    End Function

    Private Function CanStartNumericLiteral() As Boolean
        Return IsPeriod(PeekChar()) OrElse IsAmpersand(PeekChar()) OrElse IsDigit(PeekChar())
    End Function

    Private Function ReadIntegerLiteral() As Long
        Dim Value As Long = 0

        While IsDigit(PeekChar())
            Dim c As Char = MakeHalfWidth(ReadChar())
            Value *= 10
            Value += AscW(c) - AscW("0"c)
        End While

        Return Value
    End Function

    Private Function ScanDateLiteral() As Token
        Dim Start As Location = CurrentLocation
        Dim PossibleEnd As Location
        Dim Month As Integer = 0
        Dim Day As Integer = 0
        Dim Year As Integer = 0
        Dim Hour As Integer = 0
        Dim Minute As Integer = 0
        Dim Second As Integer = 0
        Dim HaveDateValue As Boolean = False
        Dim HaveTimeValue As Boolean = False
        Dim Value As Long

        Debug.Assert(CanStartDateLiteral())

        ReadChar()
        PossibleEnd = CurrentLocation
        EatWhitespace()

        ' Distinguish between date literals and the # punctuator
        If Not IsDigit(PeekChar()) Then
            Return New PunctuatorToken(TokenType.Pound, New Span(Start, PossibleEnd))
        End If

        Value = ReadIntegerLiteral()

        'LC in VBScript, it is legal to have something like #08 / 27 / 97 5:11:42pm#
        EatWhitespace()

        If IsForwardSlash(PeekChar()) OrElse IsMinus(PeekChar()) Then
            Dim Separator As Char = ReadChar()
            Dim YearStart As Location

            HaveDateValue = True
            If Value < 1 OrElse Value > 12 Then GoTo Invalid
            Month = CInt(Value)

            'LC
            EatWhitespace()

            If Not IsDigit(PeekChar()) Then GoTo Invalid
            Value = ReadIntegerLiteral()
            If Value < 1 OrElse Value > 31 Then GoTo Invalid
            Day = CInt(Value)

            'LC
            EatWhitespace()

            If PeekChar() <> Separator Then GoTo Invalid
            ReadChar()

            'LC
            EatWhitespace()

            If Not IsDigit(PeekChar()) Then GoTo Invalid
            YearStart = CurrentLocation
            Value = ReadIntegerLiteral()
            If Value < 1 OrElse Value > 9999 Then GoTo Invalid
            ' Years less than 1000 have to be four digits long to avoid y2k confusion
            'If Value < 1000 And CurrentLocation.Column - YearStart.Column <> 4 Then GoTo Invalid

            Year = CInt(Value)

            'LC 2 digit year conversion
            If CurrentLocation.Column - YearStart.Column = 2 Then
                If Year > 30 Then
                    Year += 2000
                Else
                    Year += 1900
                End If
            ElseIf CurrentLocation.Column - YearStart.Column <> 4 Then
                GoTo Invalid
            End If

            If Day > Date.DaysInMonth(Year, Month) Then GoTo Invalid

            EatWhitespace()
            If IsDigit(PeekChar()) Then
                Value = ReadIntegerLiteral()

                If Not IsColon(PeekChar()) Then GoTo Invalid
            End If
        End If

        If IsColon(PeekChar()) Then
            ReadChar()
            HaveTimeValue = True
            If Value < 0 OrElse Value > 23 Then GoTo Invalid
            Hour = CInt(Value)

            If Not IsDigit(PeekChar()) Then GoTo Invalid
            Value = ReadIntegerLiteral()
            If Value < 0 OrElse Value > 59 Then GoTo Invalid
            Minute = CInt(Value)

            If IsColon(PeekChar()) Then
                ReadChar()
                If Not IsDigit(PeekChar()) Then GoTo Invalid
                Value = ReadIntegerLiteral()
                If Value < 0 OrElse Value > 59 Then GoTo Invalid
                Second = CInt(Value)
            End If

            EatWhitespace()

            If IsA(PeekChar()) Then
                ReadChar()

                If IsM(PeekChar()) Then
                    ReadChar()
                    If Hour < 1 OrElse Hour > 12 Then
                        GoTo Invalid
                    End If
                Else
                    GoTo Invalid
                End If
            ElseIf IsP(PeekChar()) Then
                ReadChar()

                If IsM(PeekChar()) Then
                    ReadChar()
                    If Hour < 1 OrElse Hour > 12 Then
                        GoTo Invalid
                    End If

                    Hour += 12

                    If Hour = 24 Then
                        Hour = 12
                    End If
                Else
                    GoTo Invalid
                End If
            End If
        End If

        If Not IsPound(PeekChar()) Then
            GoTo Invalid
        Else
            ReadChar()
        End If

        If Not HaveTimeValue AndAlso Not HaveDateValue Then
Invalid:
            While Not IsPound(PeekChar()) AndAlso Not CanStartLineTerminator()
                ReadChar()
            End While

            If IsPound(PeekChar()) Then
                ReadChar()
            End If

            Return New ErrorToken(SyntaxErrorType.InvalidDateLiteral, SpanFrom(Start))
        End If

        If HaveDateValue Then
            If HaveTimeValue Then
                Return New DateLiteralToken(New Date(Year, Month, Day, Hour, Minute, Second), SpanFrom(Start))
            Else
                Return New DateLiteralToken(New Date(Year, Month, Day), SpanFrom(Start))
            End If
        Else
            Return New DateLiteralToken(New Date(1, 1, 1, Hour, Minute, Second), SpanFrom(Start))
        End If
    End Function

    Private Function CanStartDateLiteral() As Boolean
        Return IsPound(PeekChar())
    End Function

    ' Actually, this scans string and char literals
    Private Function ScanStringLiteral() As Token
        Dim Start As Location = CurrentLocation
        Dim s As StringBuilder = New StringBuilder()

        Debug.Assert(CanStartStringLiteral())

        ReadChar()

ContinueScan:
        While Not IsDoubleQuote(PeekChar()) AndAlso Not CanStartLineTerminator()
            s.Append(ReadChar())
        End While

        If IsDoubleQuote(PeekChar()) Then
            ReadChar()

            If IsDoubleQuote(PeekChar()) Then
                ReadChar()
                ' NOTE: We take what could be a full-width double quote and make it a half-width.
                s.Append(""""c)
                GoTo ContinueScan
            End If
        Else
            Return New ErrorToken(SyntaxErrorType.InvalidStringLiteral, SpanFrom(Start))
        End If

        If IsCharDesignator(PeekChar()) Then
            ReadChar()

            If s.Length <> 1 Then
                Return New ErrorToken(SyntaxErrorType.InvalidCharacterLiteral, SpanFrom(Start))
            Else
                Return New CharacterLiteralToken(s(0), SpanFrom(Start))
            End If
        Else
            Return New StringLiteralToken(s.ToString(), SpanFrom(Start))
        End If
    End Function

    Private Function CanStartStringLiteral() As Boolean
        Return IsDoubleQuote(PeekChar())
    End Function

    Private Function ScanIdentifier() As Token
        Dim Start As Location = CurrentLocation
        Dim Escaped As Boolean = False
        Dim TypeCharacter As TypeCharacter = TypeCharacter.None
        Dim Identifier As String
        Dim s As StringBuilder = New StringBuilder()
        Dim Type As TokenType = TokenType.Identifier
        Dim UnreservedType As TokenType = TokenType.Identifier

        Debug.Assert(CanStartIdentifier())

        If IsLeftBracket(PeekChar()) Then
            Escaped = True
            ReadChar()

            If Not CanStartNonEscapedIdentifier() Then
                While Not IsRightBracket(PeekChar()) AndAlso Not CanStartLineTerminator()
                    ReadChar()
                End While

                If IsRightBracket(PeekChar()) Then
                    ReadChar()
                End If

                Return New ErrorToken(SyntaxErrorType.InvalidEscapedIdentifier, SpanFrom(Start))
            End If
        End If

        s.Append(ReadChar())

        If IsUnderscore(s(0)) AndAlso Not IsIdentifier(PeekChar()) Then
            Dim [End] As Location = CurrentLocation

            EatWhitespace()

            ' This is a line continuation
            If CanStartLineTerminator() Then
                ScanLineTerminator(False)
                Return Nothing
            Else
                Return New ErrorToken(SyntaxErrorType.InvalidIdentifier, New Span(Start, [End]))
            End If
        End If

        While IsIdentifier(PeekChar())
            ' NOTE: We do not convert full-width to half-width here!
            s.Append(ReadChar())
        End While

        Identifier = s.ToString()

        If Escaped Then
            If IsRightBracket(PeekChar()) Then
                ReadChar()
            Else
                While Not IsRightBracket(PeekChar()) AndAlso Not CanStartLineTerminator()
                    ReadChar()
                End While

                If IsRightBracket(PeekChar()) Then
                    ReadChar()
                End If

                Return New ErrorToken(SyntaxErrorType.InvalidEscapedIdentifier, SpanFrom(Start))
            End If
        Else
            Const ValidTypeChars As TypeCharacter = _
                TypeCharacter.DecimalSymbol Or TypeCharacter.DoubleSymbol Or _
                TypeCharacter.IntegerSymbol Or TypeCharacter.LongSymbol Or _
                TypeCharacter.SingleSymbol Or TypeCharacter.StringSymbol
            Type = IdentifierToken.TokenTypeFromString(Identifier, _Version, False)

            If Type = TokenType.[REM] Then
                Return ScanComment(Start)
            End If

            UnreservedType = IdentifierToken.TokenTypeFromString(Identifier, _Version, True)
            'LC In VBScript, we do not allow type character after the identifier
            'TypeCharacter = ScanPossibleTypeCharacter(ValidTypeChars)

            If Type <> TokenType.Identifier AndAlso TypeCharacter <> TypeCharacter.None Then
                ' In VB 8.0, keywords with a type character are considered identifiers.
                If _Version > LanguageVersion.VisualBasic71 Then
                    Type = TokenType.Identifier
                Else
                    Return New ErrorToken(SyntaxErrorType.InvalidTypeCharacterOnKeyword, SpanFrom(Start))
                End If
            End If
        End If

        Return New IdentifierToken(Type, UnreservedType, Identifier, Escaped, TypeCharacter, SpanFrom(Start))
    End Function

    Private Function CanStartNonEscapedIdentifier() As Boolean
        Return CanStartNonEscapedIdentifier(PeekChar())
    End Function

    Private Shared Function CanStartNonEscapedIdentifier(ByVal c As Char) As Boolean
        Dim CharClass As UnicodeCategory = Char.GetUnicodeCategory(c)

        Return IsAlphaClass(CharClass) OrElse IsUnderscoreClass(CharClass)
    End Function

    Private Function CanStartIdentifier() As Boolean
        Return CanStartIdentifier(PeekChar())
    End Function

    Private Shared Function CanStartIdentifier(ByVal c As Char) As Boolean
        Return IsLeftBracket(c) OrElse CanStartNonEscapedIdentifier(c)
    End Function

    ' Scan a comment that begins with a tick mark
    Private Function ScanComment() As CommentToken
        Dim s As StringBuilder = New StringBuilder()
        Dim Start As Location = CurrentLocation

        Debug.Assert(CanStartComment())
        ReadChar()

        While Not CanStartLineTerminator()
            ' NOTE: We don't convert full-width to half-width here.
            s.Append(ReadChar())
        End While

        Return New CommentToken(s.ToString(), False, SpanFrom(Start))
    End Function

    ' Scan a comment that begins with REM.
    Private Function ScanComment(ByVal start As Location) As CommentToken
        Dim s As StringBuilder = New StringBuilder()

        While Not CanStartLineTerminator()
            ' NOTE: We don't convert full-width to half-width here.
            s.Append(ReadChar())
        End While

        Return New CommentToken(s.ToString(), True, SpanFrom(start))
    End Function

    ' We only check for tick mark here.
    Private Function CanStartComment() As Boolean
        Return IsSingleQuote(PeekChar())
    End Function

    Private Function ScanLineTerminator(Optional ByVal produceToken As Boolean = True) As Token
        Dim Start As Location = CurrentLocation
        Dim Token As Token = Nothing

        Debug.Assert(CanStartLineTerminator())

        If PeekChar() = ChrW(&HFFFF) Then
            Token = New EndOfStreamToken(SpanFrom(Start))
        Else
            If ReadChar() = ChrW(&HD) Then
                ' A CR/LF pair is a legal line terminator
                If PeekChar() = ChrW(&HA) Then
                    ReadChar()
                End If
            End If

            If produceToken Then
                Token = New LineTerminatorToken(SpanFrom(Start))
            End If
            _Line += 1
            _Column = 1
        End If

        Return Token
    End Function

    Private Function CanStartLineTerminator() As Boolean
        Dim c As Char = PeekChar()

        Return c = ChrW(&HD) OrElse c = ChrW(&HA) OrElse _
               c = ChrW(&H2028) OrElse c = ChrW(&H2029) OrElse _
               c = ChrW(&HFFFF)
    End Function

    Private Function EatWhitespace() As Boolean
        Dim c As Char = PeekChar()

        While c = ChrW(9) OrElse Char.GetUnicodeCategory(c) = UnicodeCategory.SpaceSeparator
            ReadChar()
            EatWhitespace = True
            c = PeekChar()
        End While
    End Function

    Private Function Read(ByVal advance As Boolean) As Token
        Dim TokenRead As Token
        Dim CurrentToken As Token

        If _Position > -1 Then
            CurrentToken = _Tokens(_Position)
        Else
            CurrentToken = Nothing
        End If

        ' If we've reached the end of the stream, just return the end of stream token again
        If CurrentToken IsNot Nothing AndAlso CurrentToken.Type = TokenType.EndOfStream Then
            Return CurrentToken
        End If

        ' If we haven't read a token yet, or if we've reached the end of the tokens that we've read
        ' so far, then we need to read a fresh token in.
        If _Position = _Tokens.Count - 1 Then
ContinueLine:
            EatWhitespace()

            If CanStartLineTerminator() Then
                TokenRead = ScanLineTerminator()
            ElseIf CanStartComment() Then
                TokenRead = ScanComment()
            ElseIf CanStartIdentifier() Then
                Dim Token As Token = ScanIdentifier()

                If Token Is Nothing Then
                    ' This was a line continuation, so skip and keep going
                    GoTo ContinueLine
                Else
                    TokenRead = Token
                End If
            ElseIf CanStartStringLiteral() Then
                TokenRead = ScanStringLiteral()
            ElseIf CanStartDateLiteral() Then
                TokenRead = ScanDateLiteral()
            ElseIf CanStartNumericLiteral() Then
                TokenRead = ScanNumericLiteral()
            Else
                Dim Start As Location = CurrentLocation
                Dim Punctuator As TokenType = PunctuatorToken.TokenTypeFromString(PeekChar())

                If Punctuator <> TokenType.None Then
                    ' CONSIDER: Only call this if we know it can start a two-character punctuator
                    TokenRead = ScanPossibleMultiCharacterPunctuator(ReadChar(), Start)
                Else
                    ReadChar()
                    TokenRead = New ErrorToken(SyntaxErrorType.InvalidCharacter, SpanFrom(Start))
                End If
            End If

            _Tokens.Add(TokenRead)
        End If

        ' Advance to the next token if we need to
        If advance Then
            _Position += 1
            Return _Tokens(_Position)
        Else
            Return _Tokens(_Position + 1)
        End If
    End Function

    ''' <summary>
    ''' Seeks backwards in the stream position to a particular token.
    ''' </summary>
    ''' <param name="token">The token to seek back to.</param>
    ''' <exception cref="System.ObjectDisposedException">Thrown when the scanner has been closed.</exception>
    ''' <exception cref="ArgumentException">Thrown when token was not produced by this scanner.</exception>
    Public Sub Seek(ByVal token As Token)
        Dim CurrentPosition As Integer
        Dim StartPosition As Integer = _Position
        Dim TokenFound As Boolean = False

        If _Disposed Then Throw New ObjectDisposedException("Scanner")

        If StartPosition = _Tokens.Count - 1 Then
            StartPosition -= 1
        End If

        For CurrentPosition = StartPosition To -1 Step -1
            If _Tokens(CurrentPosition + 1) Is token Then
                TokenFound = True
                Exit For
            End If
        Next

        If Not TokenFound Then
            Throw New ArgumentException("Token not created by this scanner.")
        Else
            _Position = CurrentPosition
        End If
    End Sub

    ''' <summary>
    ''' Whether the stream is positioned on the first token.
    ''' </summary>
    Public ReadOnly Property IsOnFirstToken() As Boolean
        Get
            Return _Position = -1
        End Get
    End Property

    ''' <summary>
    ''' Fetches the previous token in the stream.
    ''' </summary>
    ''' <returns>The previous token.</returns>
    ''' <exception cref="System.ObjectDisposedException">Thrown when the scanner has been closed.</exception>
    ''' <exception cref="System.InvalidOperationException">Thrown when the scanner is positioned on the first token.</exception>
    Public Function Previous() As Token
        If _Disposed Then Throw New ObjectDisposedException("Scanner")

        If _Position = -1 Then
            Throw New InvalidOperationException("Scanner is positioned on the first token.")
        Else
            Return _Tokens(_Position)
        End If
    End Function

    ''' <summary>
    ''' Fetches the current token without advancing the stream position.
    ''' </summary>
    ''' <returns>The current token.</returns>
    ''' <exception cref="System.ObjectDisposedException">Thrown when the scanner has been closed.</exception>
    Public Function Peek() As Token
        If _Disposed Then Throw New ObjectDisposedException("Scanner")
        Return Read(False)
    End Function

    ''' <summary>
    ''' Fetches the current token and advances the stream position.
    ''' </summary>
    ''' <returns>The current token.</returns>
    ''' <exception cref="System.ObjectDisposedException">Thrown when the scanner has been closed.</exception>
    Public Function Read() As Token
        If _Disposed Then Throw New ObjectDisposedException("Scanner")
        Return Read(True)
    End Function

    ''' <summary>
    ''' Fetches more than one token at a time from the stream.
    ''' </summary>
    ''' <param name="buffer">The array to put the tokens into.</param>
    ''' <param name="index">The location in the array to start putting the tokens into.</param>
    ''' <param name="count">The number of tokens to read.</param>
    ''' <returns>The number of tokens read.</returns>
    ''' <exception cref="System.ObjectDisposedException">Thrown when the scanner has been closed.</exception>
    ''' <exception cref="System.NullReferenceException">Thrown when the buffer is Nothing.</exception>
    ''' <exception cref="System.ArgumentException">Thrown when the index or count is invalid, or when there isn't enough room in the buffer.</exception>
    Public Function ReadBlock(ByVal buffer() As Token, ByVal index As Integer, ByVal count As Integer) As Integer
        Dim FinalCount As Integer = 0

        If buffer Is Nothing Then Throw New ArgumentNullException("buffer")
        If (index < 0) OrElse (count < 0) Then Throw New ArgumentException("Index or count cannot be negative.")
        If buffer.Length - index < count Then Throw New ArgumentException("Not enough room in buffer.")
        If _Disposed Then Throw New ObjectDisposedException("Scanner")

        While count > 0
            Dim CurrentToken As Token = Read()

            If CurrentToken.Type = TokenType.EndOfStream Then
                Return FinalCount
            End If

            buffer(FinalCount + index) = CurrentToken
            count -= 1
            FinalCount += 1
        End While

        Return FinalCount
    End Function

    ''' <summary>
    ''' Reads all of the tokens between the current position and the end of the line (or the end of the stream).
    ''' </summary>
    ''' <returns>The tokens read.</returns>
    ''' <exception cref="System.ObjectDisposedException">Thrown when the scanner has been closed.</exception>
    Public Function ReadLine() As Token()
        Dim TokenArray As List(Of Token) = New List(Of Token)()

        If _Disposed Then Throw New ObjectDisposedException("Scanner")

        While Peek().Type <> TokenType.EndOfStream And Peek().Type <> TokenType.LineTerminator
            TokenArray.Add(Read())
        End While

        Return TokenArray.ToArray()
    End Function

    ''' <summary>
    ''' Reads all the tokens between the current position and the end of the stream.
    ''' </summary>
    ''' <returns>The tokens read.</returns>
    ''' <exception cref="System.ObjectDisposedException">Thrown when the scanner has been closed.</exception>
    Public Function ReadToEnd() As Token()
        Dim TokenArray As List(Of Token) = New List(Of Token)()

        If _Disposed Then Throw New ObjectDisposedException("Scanner")

        While Peek().Type <> TokenType.EndOfStream
            TokenArray.Add(Read())
        End While

        Return TokenArray.ToArray()
    End Function

    Private Function ParseInt(ByVal literal As String) As Integer
        Dim base As Integer
        If literal.StartsWith("&"c) Then
            If literal(1) = "H"c OrElse literal(1) = "h"c Then
                base = 16
            Else
                base = 8 'Assume oct here
            End If
            literal = literal.Substring(2)
        Else
            base = 10
        End If
        Return Convert.ToInt32(literal, base)
    End Function

End Class
