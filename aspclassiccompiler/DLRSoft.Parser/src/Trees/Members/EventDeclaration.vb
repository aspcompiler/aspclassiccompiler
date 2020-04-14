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
''' A parse tree for an event declaration.
''' </summary>
Public NotInheritable Class EventDeclaration
    Inherits SignatureDeclaration

    Private ReadOnly _ImplementsList As NameCollection

    ''' <summary>
    ''' The list of implemented members.
    ''' </summary>
    Public ReadOnly Property ImplementsList() As NameCollection
        Get
            Return _ImplementsList
        End Get
    End Property

    ''' <summary>
    ''' Constructs a parse tree for an event declaration.
    ''' </summary>
    ''' <param name="attributes">The attributes for the parse tree.</param>
    ''' <param name="modifiers">The modifiers for the parse tree.</param>
    ''' <param name="keywordLocation">The location of the keyword.</param>
    ''' <param name="name">The name of the declaration.</param>
    ''' <param name="parameters">The parameters of the declaration.</param>
    ''' <param name="asLocation">The location of the 'As', if any.</param>
    ''' <param name="resultTypeAttributes">The attributes on the result type, if any.</param>
    ''' <param name="resultType">The result type, if any.</param>
    ''' <param name="implementsList">The list of implemented members.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal attributes As AttributeBlockCollection, ByVal modifiers As ModifierCollection, ByVal keywordLocation As Location, ByVal name As SimpleName, ByVal parameters As ParameterCollection, ByVal asLocation As Location, ByVal resultTypeAttributes As AttributeBlockCollection, ByVal resultType As TypeName, ByVal implementsList As NameCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.EventDeclaration, attributes, modifiers, keywordLocation, name, Nothing, parameters, asLocation, resultTypeAttributes, resultType, span, comments)

        SetParent(implementsList)
        _ImplementsList = implementsList
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)
        AddChild(childList, ImplementsList)
    End Sub
End Class