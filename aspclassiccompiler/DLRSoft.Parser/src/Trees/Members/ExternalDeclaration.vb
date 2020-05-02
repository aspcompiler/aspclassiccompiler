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
''' A parse tree for a Declare statement.
''' </summary>
Public MustInherit Class ExternalDeclaration
    Inherits SignatureDeclaration

    Private ReadOnly _CharsetLocation As Location
    Private ReadOnly _Charset As Charset
    Private ReadOnly _SubOrFunctionLocation As Location
    Private ReadOnly _LibLocation As Location
    Private ReadOnly _LibLiteral As StringLiteralExpression
    Private ReadOnly _AliasLocation As Location
    Private ReadOnly _AliasLiteral As StringLiteralExpression

    ''' <summary>
    ''' The location of 'Auto', 'Ansi' or 'Unicode', if any.
    ''' </summary>
    Public ReadOnly Property CharsetLocation() As Location
        Get
            Return _CharsetLocation
        End Get
    End Property

    ''' <summary>
    ''' The charset.
    ''' </summary>
    Public ReadOnly Property Charset() As Charset
        Get
            Return _Charset
        End Get
    End Property

    ''' <summary>
    ''' The location of 'Sub' or 'Function'.
    ''' </summary>
    Public ReadOnly Property SubOrFunctionLocation() As Location
        Get
            Return _SubOrFunctionLocation
        End Get
    End Property

    ''' <summary>
    ''' The location of 'Lib', if any.
    ''' </summary>
    Public ReadOnly Property LibLocation() As Location
        Get
            Return _LibLocation
        End Get
    End Property

    ''' <summary>
    ''' The library, if any.
    ''' </summary>
    Public ReadOnly Property LibLiteral() As StringLiteralExpression
        Get
            Return _LibLiteral
        End Get
    End Property

    ''' <summary>
    ''' The location of 'Alias', if any.
    ''' </summary>
    Public ReadOnly Property AliasLocation() As Location
        Get
            Return _AliasLocation
        End Get
    End Property

    ''' <summary>
    ''' The alias, if any.
    ''' </summary>
    Public ReadOnly Property AliasLiteral() As StringLiteralExpression
        Get
            Return _AliasLiteral
        End Get
    End Property

    Protected Sub New(ByVal type As TreeType, ByVal attributes As AttributeBlockCollection, ByVal modifiers As ModifierCollection, ByVal keywordLocation As Location, ByVal charsetLocation As Location, ByVal charset As Charset, ByVal subOrFunctionLocation As Location, ByVal name As SimpleName, ByVal libLocation As Location, ByVal libLiteral As StringLiteralExpression, ByVal aliasLocation As Location, ByVal aliasLiteral As StringLiteralExpression, ByVal parameters As ParameterCollection, ByVal asLocation As Location, ByVal resultTypeAttributes As AttributeBlockCollection, ByVal resultType As TypeName, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(type, attributes, modifiers, keywordLocation, name, Nothing, parameters, asLocation, resultTypeAttributes, resultType, span, comments)

        SetParent(libLiteral)
        SetParent(aliasLiteral)

        _CharsetLocation = charsetLocation
        _Charset = charset
        _SubOrFunctionLocation = subOrFunctionLocation
        _LibLocation = libLocation
        _LibLiteral = libLiteral
        _AliasLocation = aliasLocation
        _AliasLiteral = aliasLiteral
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)
        AddChild(childList, LibLiteral)
        AddChild(childList, AliasLiteral)
    End Sub
End Class