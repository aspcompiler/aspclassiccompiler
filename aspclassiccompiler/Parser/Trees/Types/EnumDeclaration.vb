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
''' A parse tree for an Enum declaration.
''' </summary>
Public NotInheritable Class EnumDeclaration
    Inherits BlockDeclaration

    Private ReadOnly _AsLocation As Location
    Private ReadOnly _ElementType As TypeName

    ''' <summary>
    ''' The location of the 'As', if any.
    ''' </summary>
    Public ReadOnly Property AsLocation() As Location
        Get
            Return _AsLocation
        End Get
    End Property

    ''' <summary>
    ''' The element type of the enumerated type, if any.
    ''' </summary>
    Public ReadOnly Property ElementType() As TypeName
        Get
            Return _ElementType
        End Get
    End Property

    ''' <summary>
    ''' Constructs a parse tree for an Enum declaration.
    ''' </summary>
    ''' <param name="attributes">The attributes for the parse tree.</param>
    ''' <param name="modifiers">The modifiers for the parse tree.</param>
    ''' <param name="keywordLocation">The location of the keyword.</param>
    ''' <param name="name">The name of the declaration.</param>
    ''' <param name="asLocation">The location of the 'As', if any.</param>
    ''' <param name="elementType">The element type of the enumerated type, if any.</param>
    ''' <param name="declarations">The enumerated values.</param>
    ''' <param name="endStatement">The end block declaration, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal attributes As AttributeBlockCollection, ByVal modifiers As ModifierCollection, ByVal keywordLocation As Location, ByVal name As SimpleName, ByVal asLocation As Location, ByVal elementType As TypeName, ByVal declarations As DeclarationCollection, ByVal endStatement As EndBlockDeclaration, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.EnumDeclaration, attributes, modifiers, keywordLocation, name, declarations, endStatement, span, comments)

        SetParent(elementType)

        _AsLocation = asLocation
        _ElementType = elementType
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)

        AddChild(childList, ElementType)
    End Sub
End Class