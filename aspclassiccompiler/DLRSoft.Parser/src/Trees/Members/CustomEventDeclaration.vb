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
''' A parse tree for a custom event declaration.
''' </summary>
Public NotInheritable Class CustomEventDeclaration
    Inherits SignatureDeclaration

    Private ReadOnly _CustomLocation As Location
    Private ReadOnly _ImplementsList As NameCollection
    Private ReadOnly _Accessors As DeclarationCollection
    Private ReadOnly _EndDeclaration As EndBlockDeclaration

    ''' <summary>
    ''' The location of the 'Custom' keyword.
    ''' </summary>
    Public ReadOnly Property CustomLocation() As Location
        Get
            Return _CustomLocation
        End Get
    End Property

    ''' <summary>
    ''' The list of implemented members.
    ''' </summary>
    Public ReadOnly Property ImplementsList() As NameCollection
        Get
            Return _ImplementsList
        End Get
    End Property

    ''' <summary>
    ''' The event accessors.
    ''' </summary>
    Public ReadOnly Property Accessors() As DeclarationCollection
        Get
            Return _Accessors
        End Get
    End Property

    ''' <summary>
    ''' The End Event declaration, if any.
    ''' </summary>
    Public ReadOnly Property EndDeclaration() As EndBlockDeclaration
        Get
            Return _EndDeclaration
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a custom property declaration.
    ''' </summary>
    ''' <param name="attributes">The attributes on the declaration.</param>
    ''' <param name="modifiers">The modifiers on the declaration.</param>
    ''' <param name="customLocation">The location of the 'Custom' keyword.</param>
    ''' <param name="keywordLocation">The location of the keyword.</param>
    ''' <param name="name">The name of the custom event.</param>
    ''' <param name="asLocation">The location of the 'As', if any.</param>
    ''' <param name="resultType">The result type, if any.</param>
    ''' <param name="implementsList">The implements list.</param>
    ''' <param name="accessors">The custom event accessors.</param>
    ''' <param name="endDeclaration">The End Event declaration, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal attributes As AttributeBlockCollection, ByVal modifiers As ModifierCollection, ByVal customLocation As Location, ByVal keywordLocation As Location, ByVal name As SimpleName, ByVal asLocation As Location, ByVal resultType As TypeName, ByVal implementsList As NameCollection, ByVal accessors As DeclarationCollection, ByVal endDeclaration As EndBlockDeclaration, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.CustomEventDeclaration, attributes, modifiers, keywordLocation, name, Nothing, Nothing, asLocation, Nothing, resultType, span, comments)

        SetParent(accessors)
        SetParent(endDeclaration)
        SetParent(implementsList)

        _CustomLocation = customLocation
        _ImplementsList = implementsList
        _Accessors = accessors
        _EndDeclaration = endDeclaration
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)
        AddChild(childList, ImplementsList)
        AddChild(childList, Accessors)
        AddChild(childList, EndDeclaration)
    End Sub
End Class