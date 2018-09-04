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
''' An identifier.
''' </summary>
Public NotInheritable Class IdentifierToken
    Inherits Token

    Private Structure Keyword
        Public ReadOnly Versions As LanguageVersion
        Public ReadOnly ReservedVersions As LanguageVersion
        Public ReadOnly TokenType As TokenType

        Public Sub New(ByVal Versions As LanguageVersion, ByVal ReservedVersions As LanguageVersion, ByVal TokenType As TokenType)
            Me.Versions = Versions
            Me.ReservedVersions = ReservedVersions
            Me.TokenType = TokenType
        End Sub
    End Structure

    Private Shared KeywordTable As Dictionary(Of String, Keyword)

    Private Shared Sub AddKeyword(ByVal table As Dictionary(Of String, Keyword), ByVal name As String, ByVal keyword As Keyword)
        table.Add(name, keyword)
        table.Add(Scanner.MakeFullWidth(name), keyword)
    End Sub

    ' Returns the token type of the string.
    Friend Shared Function TokenTypeFromString(ByVal s As String, ByVal Version As LanguageVersion, ByVal IncludeUnreserved As Boolean) As TokenType
        If KeywordTable Is Nothing Then
            Dim Table As New Dictionary(Of String, Keyword)(StringComparer.InvariantCultureIgnoreCase)

            ' NOTE: These have to be in the same order as the enum!
            AddKeyword(Table, "AddHandler", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.AddHandler))
            AddKeyword(Table, "AddressOf", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.AddressOf))
            AddKeyword(Table, "Alias", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Alias))
            AddKeyword(Table, "And", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.And))
            AddKeyword(Table, "AndAlso", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.AndAlso))
            AddKeyword(Table, "Ansi", New Keyword(LanguageVersion.All, LanguageVersion.VisualBasic71, TokenType.Ansi))
            AddKeyword(Table, "As", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.As))
            AddKeyword(Table, "Assembly", New Keyword(LanguageVersion.All, LanguageVersion.VisualBasic71, TokenType.Assembly))
            AddKeyword(Table, "Auto", New Keyword(LanguageVersion.All, LanguageVersion.VisualBasic71, TokenType.Auto))
            AddKeyword(Table, "Binary", New Keyword(LanguageVersion.All, LanguageVersion.None, TokenType.Binary))
            AddKeyword(Table, "Boolean", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Boolean))
            AddKeyword(Table, "ByRef", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.ByRef))
            AddKeyword(Table, "Byte", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Byte))
            AddKeyword(Table, "ByVal", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.ByVal))
            AddKeyword(Table, "Call", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Call))
            AddKeyword(Table, "Case", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Case))
            AddKeyword(Table, "Catch", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Catch))
            'AddKeyword(Table, "CBool", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CBool))
            'AddKeyword(Table, "CByte", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CByte))
            'AddKeyword(Table, "CChar", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CChar))
            'AddKeyword(Table, "CDate", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CDate))
            'AddKeyword(Table, "CDbl", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CDbl))
            'AddKeyword(Table, "CDec", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CDec))
            AddKeyword(Table, "Char", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Char))
            'AddKeyword(Table, "CInt", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CInt))
            AddKeyword(Table, "Class", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Class))
            'AddKeyword(Table, "CLng", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CLng))
            'AddKeyword(Table, "CObj", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CObj))
            AddKeyword(Table, "Compare", New Keyword(LanguageVersion.All, LanguageVersion.None, TokenType.Compare))
            AddKeyword(Table, "Const", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Const))
            AddKeyword(Table, "Continue", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.Continue))
            'AddKeyword(Table, "CSByte", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.CSByte))
            'AddKeyword(Table, "CShort", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CShort))
            'AddKeyword(Table, "CSng", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CSng))
            'AddKeyword(Table, "CStr", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CStr))
            'AddKeyword(Table, "CType", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CType))
            'AddKeyword(Table, "CUInt", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.CUInt))
            'AddKeyword(Table, "CULng", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.CULng))
            'AddKeyword(Table, "CUShort", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.CUShort))
            AddKeyword(Table, "Custom", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.None, TokenType.Custom))
            AddKeyword(Table, "Date", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Date))
            AddKeyword(Table, "Decimal", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Decimal))
            AddKeyword(Table, "Declare", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Declare))
            AddKeyword(Table, "Default", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Default))
            AddKeyword(Table, "Delegate", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Delegate))
            AddKeyword(Table, "Dim", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Dim))
            AddKeyword(Table, "DirectCast", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.DirectCast))
            AddKeyword(Table, "Do", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Do))
            AddKeyword(Table, "Double", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Double))
            AddKeyword(Table, "Each", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Each))
            AddKeyword(Table, "Else", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Else))
            AddKeyword(Table, "ElseIf", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.ElseIf))
            AddKeyword(Table, "End", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.End))
            AddKeyword(Table, "EndIf", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.EndIf))
            AddKeyword(Table, "Enum", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Enum))
            AddKeyword(Table, "Erase", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Erase))
            AddKeyword(Table, "Error", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Error))
            AddKeyword(Table, "Event", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Event))
            AddKeyword(Table, "Exit", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Exit))
            AddKeyword(Table, "Explicit", New Keyword(LanguageVersion.All, LanguageVersion.None, TokenType.Explicit))
            AddKeyword(Table, "ExternalChecksum", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.None, TokenType.ExternalChecksum))
            AddKeyword(Table, "ExternalSource", New Keyword(LanguageVersion.All, LanguageVersion.None, TokenType.ExternalSource))
            AddKeyword(Table, "False", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.False))
            AddKeyword(Table, "Finally", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Finally))
            AddKeyword(Table, "For", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.For))
            AddKeyword(Table, "Friend", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Friend))
            AddKeyword(Table, "Function", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Function))
            AddKeyword(Table, "Get", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Get))
            AddKeyword(Table, "GetType", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.GetType))
            AddKeyword(Table, "Global", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.Global))
            AddKeyword(Table, "GoSub", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.GoSub))
            AddKeyword(Table, "GoTo", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.GoTo))
            AddKeyword(Table, "Handles", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Handles))
            AddKeyword(Table, "If", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.If))
            AddKeyword(Table, "Implements", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Implements))
            AddKeyword(Table, "Imports", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Imports))
            AddKeyword(Table, "In", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.In))
            AddKeyword(Table, "Inherits", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Inherits))
            AddKeyword(Table, "Integer", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Integer))
            AddKeyword(Table, "Interface", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Interface))
            AddKeyword(Table, "Is", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Is))
            AddKeyword(Table, "IsFalse", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.None, TokenType.IsFalse))
            AddKeyword(Table, "IsNot", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.IsNot))
            AddKeyword(Table, "IsTrue", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.None, TokenType.IsTrue))
            AddKeyword(Table, "Let", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Let))
            AddKeyword(Table, "Lib", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Lib))
            AddKeyword(Table, "Like", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Like))
            AddKeyword(Table, "Long", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Long))
            AddKeyword(Table, "Loop", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Loop))
            AddKeyword(Table, "Me", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Me))
            AddKeyword(Table, "Mid", New Keyword(LanguageVersion.All, LanguageVersion.None, TokenType.Mid))
            AddKeyword(Table, "Mod", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Mod))
            AddKeyword(Table, "Module", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Module))
            AddKeyword(Table, "MustInherit", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.MustInherit))
            AddKeyword(Table, "MustOverride", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.MustOverride))
            AddKeyword(Table, "MyBase", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.MyBase))
            AddKeyword(Table, "MyClass", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.MyClass))
            AddKeyword(Table, "Namespace", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Namespace))
            AddKeyword(Table, "Narrowing", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.Narrowing))
            AddKeyword(Table, "New", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.New))
            AddKeyword(Table, "Next", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Next))
            AddKeyword(Table, "Not", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Not))
            AddKeyword(Table, "Nothing", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Nothing))
            AddKeyword(Table, "NotInheritable", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.NotInheritable))
            AddKeyword(Table, "NotOverridable", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.NotOverridable))
            AddKeyword(Table, "Object", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Object))
            AddKeyword(Table, "Of", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.Of))
            AddKeyword(Table, "Off", New Keyword(LanguageVersion.All, LanguageVersion.None, TokenType.Off))
            AddKeyword(Table, "On", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.On))
            AddKeyword(Table, "Operator", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.Operator))
            AddKeyword(Table, "Option", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Option))
            AddKeyword(Table, "Optional", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Optional))
            AddKeyword(Table, "Or", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Or))
            AddKeyword(Table, "OrElse", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.OrElse))
            AddKeyword(Table, "Overloads", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Overloads))
            AddKeyword(Table, "Overridable", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Overridable))
            AddKeyword(Table, "Overrides", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Overrides))
            AddKeyword(Table, "ParamArray", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.ParamArray))
            AddKeyword(Table, "Partial", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.Partial))
            AddKeyword(Table, "Preserve", New Keyword(LanguageVersion.All, LanguageVersion.VisualBasic71, TokenType.Preserve))
            AddKeyword(Table, "Private", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Private))
            AddKeyword(Table, "Property", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Property))
            AddKeyword(Table, "Protected", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Protected))
            AddKeyword(Table, "Public", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Public))
            AddKeyword(Table, "RaiseEvent", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.RaiseEvent))
            AddKeyword(Table, "ReadOnly", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.ReadOnly))
            AddKeyword(Table, "ReDim", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.ReDim))
            AddKeyword(Table, "Region", New Keyword(LanguageVersion.All, LanguageVersion.None, TokenType.Region))
            AddKeyword(Table, "REM", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.[REM]))
            AddKeyword(Table, "RemoveHandler", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.RemoveHandler))
            AddKeyword(Table, "Resume", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Resume))
            AddKeyword(Table, "Return", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Return))
            AddKeyword(Table, "SByte", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.SByte))
            AddKeyword(Table, "Select", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Select))
            AddKeyword(Table, "Set", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Set))
            AddKeyword(Table, "Shadows", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Shadows))
            AddKeyword(Table, "Shared", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Shared))
            AddKeyword(Table, "Short", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Short))
            AddKeyword(Table, "Single", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Single))
            AddKeyword(Table, "Static", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Static))
            AddKeyword(Table, "Step", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Step))
            AddKeyword(Table, "Stop", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Stop))
            AddKeyword(Table, "Strict", New Keyword(LanguageVersion.All, LanguageVersion.None, TokenType.Strict))
            AddKeyword(Table, "String", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.String))
            AddKeyword(Table, "Structure", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Structure))
            AddKeyword(Table, "Sub", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Sub))
            AddKeyword(Table, "SyncLock", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.SyncLock))
            AddKeyword(Table, "Text", New Keyword(LanguageVersion.All, LanguageVersion.None, TokenType.Text))
            AddKeyword(Table, "Then", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Then))
            AddKeyword(Table, "Throw", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Throw))
            AddKeyword(Table, "To", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.To))
            AddKeyword(Table, "True", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.True))
            AddKeyword(Table, "Try", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Try))
            AddKeyword(Table, "TryCast", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.TryCast))
            AddKeyword(Table, "TypeOf", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.TypeOf))
            AddKeyword(Table, "UInteger", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.UInteger))
            AddKeyword(Table, "ULong", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.ULong))
            AddKeyword(Table, "UShort", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.UShort))
            AddKeyword(Table, "Using", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.Using))
            AddKeyword(Table, "Unicode", New Keyword(LanguageVersion.All, LanguageVersion.VisualBasic71, TokenType.Unicode))
            AddKeyword(Table, "Until", New Keyword(LanguageVersion.All, LanguageVersion.VisualBasic71, TokenType.Until))
            AddKeyword(Table, "Variant", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Variant))
            AddKeyword(Table, "Wend", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Wend))
            AddKeyword(Table, "When", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.When))
            AddKeyword(Table, "While", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.While))
            AddKeyword(Table, "Widening", New Keyword(LanguageVersion.VisualBasic80, LanguageVersion.VisualBasic80, TokenType.Widening))
            AddKeyword(Table, "With", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.With))
            AddKeyword(Table, "WithEvents", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.WithEvents))
            AddKeyword(Table, "WriteOnly", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.WriteOnly))
            AddKeyword(Table, "Xor", New Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Xor))

            KeywordTable = Table
        End If

        If KeywordTable.ContainsKey(s) Then
            Dim Keyword As Keyword = KeywordTable(s)

            If ((Keyword.Versions And Version) = Version) AndAlso _
               (IncludeUnreserved OrElse ((Keyword.ReservedVersions And Version) = Version)) Then
                Return Keyword.TokenType
            End If
        End If

        Return TokenType.Identifier
    End Function

    ''' <summary>
    ''' Determines if a token type is a keyword.
    ''' </summary>
    ''' <param name="type">The token type to check.</param>
    ''' <returns>True if the token type is a keyword, False otherwise.</returns>
    Public Shared Function IsKeyword(ByVal type As TokenType) As Boolean
        Return type >= TokenType.AddHandler AndAlso type <= TokenType.Xor
    End Function

    Public Overrides Function AsUnreservedKeyword() As TokenType
        Return _UnreservedType
    End Function

    Private ReadOnly _Identifier As String
    Private ReadOnly _UnreservedType As TokenType
    Private ReadOnly _Escaped As Boolean                  ' Whether the identifier was escaped (i.e. [a])
    Private ReadOnly _TypeCharacter As TypeCharacter      ' The type character that followed, if any

    ''' <summary>
    ''' The identifier name.
    ''' </summary>
    Public ReadOnly Property Identifier() As String
        Get
            Return _Identifier
        End Get
    End Property

    ''' <summary>
    ''' Whether the identifier is escaped.
    ''' </summary>
    Public ReadOnly Property Escaped() As Boolean
        Get
            Return _Escaped
        End Get
    End Property

    ''' <summary>
    ''' The type character of the identifier.
    ''' </summary>
    Public ReadOnly Property TypeCharacter() As TypeCharacter
        Get
            Return _TypeCharacter
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new identifier token.
    ''' </summary>
    ''' <param name="type">The token type of the identifier.</param>
    ''' <param name="unreservedType">The unreserved token type of the identifier.</param>
    ''' <param name="identifier">The text of the identifier</param>
    ''' <param name="escaped">Whether the identifier is escaped.</param>
    ''' <param name="typeCharacter">The type character of the identifier.</param>
    ''' <param name="span">The location of the identifier.</param>
    Public Sub New(ByVal type As TokenType, ByVal unreservedType As TokenType, ByVal identifier As String, ByVal escaped As Boolean, ByVal typeCharacter As TypeCharacter, ByVal span As Span)
        MyBase.New(type, span)

        If type <> TokenType.Identifier AndAlso Not IsKeyword(type) Then
            Throw New ArgumentOutOfRangeException("type")
        End If

        If unreservedType <> TokenType.Identifier AndAlso Not IsKeyword(unreservedType) Then
            Throw New ArgumentOutOfRangeException("unreservedType")
        End If

        If identifier Is Nothing OrElse identifier = "" Then
            Throw New ArgumentException("Identifier cannot be empty.", "identifier")
        End If

        If typeCharacter <> typeCharacter.None AndAlso typeCharacter <> typeCharacter.DecimalSymbol AndAlso _
           typeCharacter <> typeCharacter.DoubleSymbol AndAlso typeCharacter <> typeCharacter.IntegerSymbol AndAlso _
           typeCharacter <> typeCharacter.LongSymbol AndAlso typeCharacter <> typeCharacter.SingleSymbol AndAlso _
           typeCharacter <> typeCharacter.StringSymbol Then
            Throw New ArgumentOutOfRangeException("typeCharacter")
        End If

        If typeCharacter <> typeCharacter.None AndAlso escaped Then
            Throw New ArgumentException("Escaped identifiers cannot have type characters.")
        End If

        _UnreservedType = unreservedType
        _Identifier = identifier
        _Escaped = escaped
        _TypeCharacter = typeCharacter
    End Sub
End Class