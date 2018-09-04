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
''' A parse tree for a namespace declaration.
''' </summary>
Public NotInheritable Class NamespaceDeclaration
    Inherits ModifiedDeclaration

    Private ReadOnly _NamespaceLocation As Location
    Private ReadOnly _Name As Name
    Private ReadOnly _Declarations As DeclarationCollection
    Private ReadOnly _EndDeclaration As EndBlockDeclaration

    ''' <summary>
    ''' The location of 'Namespace'.
    ''' </summary>
    Public ReadOnly Property NamespaceLocation() As Location
        Get
            Return _NamespaceLocation
        End Get
    End Property

    ''' <summary>
    ''' The name of the namespace.
    ''' </summary>
    Public ReadOnly Property Name() As Name
        Get
            Return _Name
        End Get
    End Property

    ''' <summary>
    ''' The declarations in the namespace.
    ''' </summary>
    Public ReadOnly Property Declarations() As DeclarationCollection
        Get
            Return _Declarations
        End Get
    End Property

    ''' <summary>
    ''' The End Namespace declaration, if any.
    ''' </summary>
    Public ReadOnly Property EndDeclaration() As EndBlockDeclaration
        Get
            Return _EndDeclaration
        End Get
    End Property

    ''' <summary>
    ''' Constructs a parse tree for a namespace declaration.
    ''' </summary>
    ''' <param name="attributes">The attributes on the declaration.</param>
    ''' <param name="modifiers">The modifiers on the declaration.</param>
    ''' <param name="namespaceLocation">The location of 'Namespace'.</param>
    ''' <param name="name">The name of the namespace.</param>
    ''' <param name="declarations">The declarations in the namespace.</param>
    ''' <param name="endDeclaration">The End Namespace statement, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal attributes As AttributeBlockCollection, ByVal modifiers As ModifierCollection, ByVal namespaceLocation As Location, ByVal name As Name, ByVal declarations As DeclarationCollection, ByVal endDeclaration As EndBlockDeclaration, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.NamespaceDeclaration, attributes, modifiers, span, comments)

        If name Is Nothing Then
            Throw New ArgumentNullException("name")
        End If

        SetParent(name)
        SetParent(declarations)
        SetParent(endDeclaration)

        _NamespaceLocation = namespaceLocation
        _Name = name
        _Declarations = declarations
        _EndDeclaration = endDeclaration
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)

        AddChild(childList, Name)
        AddChild(childList, Declarations)
        AddChild(childList, EndDeclaration)
    End Sub
End Class