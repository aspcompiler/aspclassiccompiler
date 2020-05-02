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
''' A parse tree for a possibly generic block declaration.
''' </summary>
Public MustInherit Class GenericBlockDeclaration
    Inherits BlockDeclaration

    Private ReadOnly _TypeParameters As TypeParameterCollection

    ''' <summary>
    ''' The type parameters of the type, if any.
    ''' </summary>
    Public ReadOnly Property TypeParameters() As TypeParameterCollection
        Get
            Return _TypeParameters
        End Get
    End Property

    Protected Sub New(ByVal type As TreeType, ByVal attributes As AttributeBlockCollection, ByVal modifiers As ModifierCollection, ByVal keywordLocation As Location, ByVal name As SimpleName, ByVal typeParameters As TypeParameterCollection, ByVal declarations As DeclarationCollection, ByVal endStatement As EndBlockDeclaration, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(type, attributes, modifiers, keywordLocation, name, declarations, endStatement, span, comments)

        Debug.Assert(type = TreeType.ClassDeclaration OrElse type = TreeType.InterfaceDeclaration OrElse type = TreeType.StructureDeclaration)

        SetParent(typeParameters)
        _TypeParameters = typeParameters
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)

        AddChild(childList, TypeParameters)
    End Sub
End Class