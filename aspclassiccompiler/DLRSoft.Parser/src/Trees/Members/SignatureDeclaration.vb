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
''' A parse tree for a declaration with a signature.
''' </summary>
Public MustInherit Class SignatureDeclaration
    Inherits ModifiedDeclaration

    Private ReadOnly _KeywordLocation As Location
    Private ReadOnly _Name As SimpleName
    Private ReadOnly _TypeParameters As TypeParameterCollection
    Private ReadOnly _Parameters As ParameterCollection
    Private ReadOnly _AsLocation As Location
    Private ReadOnly _ResultTypeAttributes As AttributeBlockCollection
    Private ReadOnly _ResultType As TypeName

    ''' <summary>
    ''' The location of the declaration's keyword.
    ''' </summary>
    Public ReadOnly Property KeywordLocation() As Location
        Get
            Return _KeywordLocation
        End Get
    End Property

    ''' <summary>
    ''' The name of the declaration.
    ''' </summary>
    Public ReadOnly Property Name() As SimpleName
        Get
            Return _Name
        End Get
    End Property

    ''' <summary>
    ''' The type parameters on the declaration, if any.
    ''' </summary>
    Public ReadOnly Property TypeParameters() As TypeParameterCollection
        Get
            Return _TypeParameters
        End Get
    End Property

    ''' <summary>
    ''' The parameters of the declaration.
    ''' </summary>
    Public ReadOnly Property Parameters() As ParameterCollection
        Get
            Return _Parameters
        End Get
    End Property

    ''' <summary>
    ''' The location of the 'As', if any.
    ''' </summary>
    Public ReadOnly Property AsLocation() As Location
        Get
            Return _AsLocation
        End Get
    End Property

    ''' <summary>
    ''' The result type attributes, if any.
    ''' </summary>
    Public ReadOnly Property ResultTypeAttributes() As AttributeBlockCollection
        Get
            Return _ResultTypeAttributes
        End Get
    End Property

    ''' <summary>
    ''' The result type, if any.
    ''' </summary>
    Public ReadOnly Property ResultType() As TypeName
        Get
            Return _ResultType
        End Get
    End Property

    Protected Sub New(ByVal type As TreeType, ByVal attributes As AttributeBlockCollection, ByVal modifiers As ModifierCollection, ByVal keywordLocation As Location, ByVal name As SimpleName, ByVal typeParameters As TypeParameterCollection, ByVal parameters As ParameterCollection, ByVal asLocation As Location, ByVal resultTypeAttributes As AttributeBlockCollection, ByVal resultType As TypeName, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(type, attributes, modifiers, span, comments)

        SetParent(name)
        SetParent(typeParameters)
        SetParent(parameters)
        SetParent(resultType)
        SetParent(resultTypeAttributes)

        _KeywordLocation = keywordLocation
        _Name = name
        _TypeParameters = typeParameters
        _Parameters = parameters
        _AsLocation = asLocation
        _ResultTypeAttributes = resultTypeAttributes
        _ResultType = resultType
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)
        AddChild(childList, Name)
        AddChild(childList, TypeParameters)
        AddChild(childList, Parameters)
        AddChild(childList, ResultTypeAttributes)
        AddChild(childList, ResultType)
    End Sub
End Class