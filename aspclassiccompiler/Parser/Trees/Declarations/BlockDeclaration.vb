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
''' A parse tree for a block declaration.
''' </summary>
Public MustInherit Class BlockDeclaration
    Inherits ModifiedDeclaration

    Private ReadOnly _KeywordLocation As Location
    Private ReadOnly _Name As SimpleName
    Private ReadOnly _Declarations As DeclarationCollection
    Private ReadOnly _EndDeclaration As EndBlockDeclaration

    ''' <summary>
    ''' The location of the keyword.
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
    ''' The declarations in the block.
    ''' </summary>
    Public ReadOnly Property Declarations() As DeclarationCollection
        Get
            Return _Declarations
        End Get
    End Property

    ''' <summary>
    ''' The End statement for the block.
    ''' </summary>
    Public ReadOnly Property EndDeclaration() As EndBlockDeclaration
        Get
            Return _EndDeclaration
        End Get
    End Property

    Protected Sub New(ByVal type As TreeType, ByVal attributes As AttributeBlockCollection, ByVal modifiers As ModifierCollection, ByVal keywordLocation As Location, ByVal name As SimpleName, ByVal declarations As DeclarationCollection, ByVal endDeclaration As EndBlockDeclaration, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(type, attributes, modifiers, span, comments)

        Debug.Assert(type = TreeType.ClassDeclaration OrElse type = TreeType.ModuleDeclaration OrElse _
                     type = TreeType.InterfaceDeclaration OrElse type = TreeType.StructureDeclaration OrElse _
                     type = TreeType.EnumDeclaration)

        If name Is Nothing Then
            Throw New ArgumentNullException("name")
        End If

        SetParent(name)
        SetParent(declarations)
        SetParent(endDeclaration)

        _KeywordLocation = keywordLocation
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