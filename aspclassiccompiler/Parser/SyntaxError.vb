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
''' A syntax error.
''' </summary>
Public NotInheritable Class SyntaxError
    Private ReadOnly _Type As SyntaxErrorType
    Private ReadOnly _Span As Span

    ''' <summary>
    ''' The type of the syntax error.
    ''' </summary>
    Public ReadOnly Property Type() As SyntaxErrorType
        Get
            Return _Type
        End Get
    End Property

    ''' <summary>
    ''' The location of the syntax error.
    ''' </summary>
    Public ReadOnly Property Span() As Span
        Get
            Return _Span
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new syntax error.
    ''' </summary>
    ''' <param name="type">The type of the syntax error.</param>
    ''' <param name="span">The location of the syntax error.</param>
    Public Sub New(ByVal type As SyntaxErrorType, ByVal span As Span)
        Debug.Assert(System.Enum.IsDefined(GetType(SyntaxErrorType), type))
        _Type = type
        _Span = span
    End Sub

    Public Overrides Function ToString() As String
        Static Messages() As String

        If Messages Is Nothing Then
            Dim Strings() As String

            Strings = New String() { _
                "", _
                "Invalid escaped identifier.", _
                "Invalid identifier.", _
                "Cannot put a type character on a keyword.", _
                "Invalid character.", _
                "Invalid character literal.", _
                "Invalid string literal.", _
                "Invalid date literal.", _
                "Invalid floating point literal.", _
                "Invalid integer literal.", _
                "Invalid decimal literal.", _
                "Syntax error.", _
                "Expected ','.", _
                "Expected '('.", _
                "Expected ')'.", _
                "Expected '='.", _
                "Expected 'As'.", _
                "Expected '}'.", _
                "Expected '.'.", _
                "Expected '-'.", _
                "Expected 'Is'.", _
                "Expected '>'.", _
                "Type expected.", _
                "Expected identifier.", _
                "Invalid use of keyword.", _
                "Bounds can be specified only for the top-level array when initializing an array of arrays.", _
                "Array bounds cannot appear in type specifiers.", _
                "Expected expression.", _
                "Comma, ')', or a valid expression continuation expected.", _
                "Expected named argument.", _
                "MyBase must be followed by a '.'.", _
                "MyClass must be followed by a '.'.", _
                "Exit must be followed by Do, For, While, Select, Sub, Function, Property or Try.", _
                "Expected 'Next'.", _
                "Expected 'Resume' or 'GoTo'.", _
                "Expected 'Error'.", _
                "Type character does not match declared data type String.", _
                "Comma, '}', or a valid expression continuation expected.", _
                "Expected one of 'Dim', 'Const', 'Public', 'Private', 'Protected', 'Friend', 'Shadows', 'ReadOnly' or 'Shared'.", _
                "End of statement expected.", _
                "'Do' must end with a matching 'Loop'.", _
                "'While' must end with a matching 'End While'.", _
                "'Select' must end with a matching 'End Select'.", _
                "'SyncLock' must end with a matching 'End SyncLock'.", _
                "'With' must end with a matching 'End With'.", _
                "'If' must end with a matching 'End If'.", _
                "'Try' must end with a matching 'End Try'.", _
                "'Sub' must end with a matching 'End Sub'.", _
                "'Function' must end with a matching 'End Function'.", _
                "'Property' must end with a matching 'End Property'.", _
                "'Get' must end with a matching 'End Get'.", _
                "'Set' must end with a matching 'End Set'.", _
                "'Class' must end with a matching 'End Class'.", _
                "'Structure' must end with a matching 'End Structure'.", _
                "'Module' must end with a matching 'End Module'.", _
                "'Interface' must end with a matching 'End Interface'.", _
                "'Enum' must end with a matching 'End Enum'.", _
                "'Namespace' must end with a matching 'End Namespace'.", _
                "'Loop' cannot have a condition if matching 'Do' has one.", _
                "'Loop' must be preceded by a matching 'Do'.", _
                "'Next' must be preceded by a matching 'For' or 'For Each'.", _
                "'End While' must be preceded by a matching 'While'.", _
                "'End Select' must be preceded by a matching 'Select'.", _
                "'End SyncLock' must be preceded by a matching 'SyncLock'.", _
                "'End If' must be preceded by a matching 'If'.", _
                "'End Try' must be preceded by a matching 'Try'.", _
                "'End With' must be preceded by a matching 'With'.", _
                "'Catch' cannot appear outside a 'Try' statement.", _
                "'Finally' cannot appear outside a 'Try' statement.", _
                "'Catch' cannot appear after 'Finally' within a 'Try' statement.", _
                "'Finally' can only appear once in a 'Try' statement.", _
                "'Case' must be preceded by a matching 'Select'.", _
                "'Case' cannot appear after 'Case Else' within a 'Select' statement.", _
                "'Case Else' can only appear once in a 'Select' statement.", _
                "'Case Else' must be preceded by a matching 'Select'.", _
                "'End Sub' must be preceded by a matching 'Sub'.", _
                "'End Function' must be preceded by a matching 'Function'.", _
                "'End Property' must be preceded by a matching 'Property'.", _
                "'End Get' must be preceded by a matching 'Get'.", _
                "'End Set' must be preceded by a matching 'Set'.", _
                "'End Class' must be preceded by a matching 'Class'.", _
                "'End Structure' must be preceded by a matching 'Structure'.", _
                "'End Module' must be preceded by a matching 'Module'.", _
                "'End Interface' must be preceded by a matching 'Interface'.", _
                "'End Enum' must be preceded by a matching 'Enum'.", _
                "'End Namespace' must be preceded by a matching 'Namespace'.", _
                "Statements and labels are not valid between 'Select Case' and first 'Case'.", _
                "'ElseIf' cannot appear after 'Else' within an 'If' statement.", _
                "'ElseIf' must be preceded by a matching 'If'.", _
                "'Else' can only appear once in an 'If' statement.", _
                "'Else' must be preceded by a matching 'If'.", _
                "Statement cannot end a block outside of a line 'If' statement.", _
                "Attribute of this type is not allowed here.", _
                "Modifier cannot be specified twice.", _
                "Modifier is not valid on this declaration type.", _
                "Can only specify one of 'Dim', 'Static' or 'Const'.", _
                "Events cannot have a return type.", _
                "Comma or ')' expected.", _
                "Method declaration statements must be the first statement on a logical line.", _
                "First statement of a method body cannot be on the same line as the method declaration.", _
                "'End Sub' must be the first statement on a line.", _
                "'End Function' must be the first statement on a line.", _
                "'End Get' must be the first statement on a line.", _
                "'End Set' must be the first statement on a line.", _
                "'Sub' or 'Function' expected.", _
                "String constant expected.", _
                "'Lib' expected.", _
                "Declaration cannot appear within a Property declaration.", _
                "Declaration cannot appear within a Class declaration.", _
                "Declaration cannot appear within a Structure declaration.", _
                "Declaration cannot appear within a Module declaration.", _
                "Declaration cannot appear within an Interface declaration.", _
                "Declaration cannot appear within an Enum declaration.", _
                "Declaration cannot appear within a Namespace declaration.", _
                "Specifiers and attributes are not valid on this statement.", _
                "Specifiers and attributes are not valid on a Namespace declaration.", _
                "Specifiers and attributes are not valid on an Imports declaration.", _
                "Specifiers and attributes are not valid on an Option declaration.", _
                "Inherits' can only specify one class.", _
                "'Inherits' statements must precede all declarations.", _
                "'Implements' statements must follow any 'Inherits' statement and precede all declarations in a class.", _
                "Enum must contain at least one member.", _
                "'Option Explicit' can be followed only by 'On' or 'Off'.", _
                "'Option Strict' can be followed only by 'On' or 'Off'.", _
                "'Option Compare' must be followed by 'Text' or 'Binary'.", _
                "'Option' must be followed by 'Compare', 'Explicit', or 'Strict'.", _
                "'Option' statements must precede any declarations or 'Imports' statements.", _
                "'Imports' statements must precede any declarations.", _
                "Assembly or Module attribute statements must precede any declarations in a file.", _
                "'End' statement not valid.", _
                "Expected relational operator.", _
                "'If', 'ElseIf', 'Else', 'End If', or 'Const' expected.", _
                "Expected integer literal.", _
                "'#ExternalSource' statements cannot be nested.", _
                "'ExternalSource', 'Region' or 'If' expected.", _
                "'#End ExternalSource' must be preceded by a matching '#ExternalSource'.", _
                "'#ExternalSource' must end with a matching '#End ExternalSource'.", _
                "'#End Region' must be preceded by a matching '#Region'.", _
                "'#Region' must end with a matching '#End Region'.", _
                "'#Region' and '#End Region' statements are not valid within method bodies.", _
                "Conversions to and from 'String' cannot occur in a conditional compilation expression.", _
                "Conversion is not valid in a conditional compilation expression.", _
                "Expression is not valid in a conditional compilation expression.", _
                "Operator is not valid for these types in a conditional compilation expression.", _
                "'#If' must end with a matching '#End If'.", _
                "'#End If' must be preceded by a matching '#If'.", _
                "'#ElseIf' cannot appear after '#Else' within an '#If' statement.", _
                "'#ElseIf' must be preceded by a matching '#If'.", _
                "'#Else' can only appear once in an '#If' statement.", _
                "'#Else' must be preceded by a matching '#If'.", _
                "'Global' not allowed in this context; identifier expected.", _
                "Modules cannot be generic.", _
                "Expected 'Of'.", _
                "Operator declaration must be one of:  +, -, *, \, /, ^, &, Like, Mod, And, Or, Xor, Not, <<, >>, =, <>, <, <=, >, >=, CType, IsTrue, IsFalse.", _
                "'Operator' must end with a matching 'End Operator'.", _
                "'End Operator' must be preceded by a matching 'Operator'.", _
                "'End Operator' must be the first statement on a line.", _
                "Properties cannot be generic.", _
                "Constructors cannot be generic.", _
                "Operators cannot be generic.", _
                "Global must be followed by a '.'.", _
                "Continue must be followed by Do, For, or While.", _
                "'Using' must end with a matching 'End Using'.", _
                "Custom 'Event' must end with a matching 'End Event'.", _
                "'AddHandler' must end with a matching 'End AddHandler'.", _
                "'RemoveHandler' must end with a matching 'End RemoveHandler'.", _
                "'RaiseEvent' must end with a matching 'End RaiseEvent'.", _
                "'End Using' must be preceded by a matching 'Using'.", _
                "'End Event' must be preceded by a matching custom 'Event'.", _
                "'End AddHandler' must be preceded by a matching 'AddHandler'.", _
                "'End RemoveHandler' must be preceded by a matching 'RemoveHandler'.", _
                "'End RaiseEvent' must be preceded by a matching 'RaiseEvent'.", _
                "'End AddHandler' must be the first statement on a line.", _
                "'End RemoveHandler' must be the first statement on a line.", _
                "'End RaiseEvent' must be the first statement on a line." _
                }

            Messages = Strings
        End If

        Return "error " & Type & " " & Span.ToString() & ": " & Messages(Type)
    End Function
End Class
