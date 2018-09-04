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
''' A parse tree for a delegate declaration.
''' </summary>
Public MustInherit Class DelegateDeclaration
    Inherits SignatureDeclaration

    Private ReadOnly _SubOrFunctionLocation As Location

    ''' <summary>
    ''' The location of 'Sub' or 'Function'.
    ''' </summary>
    Public ReadOnly Property SubOrFunctionLocation() As Location
        Get
            Return _SubOrFunctionLocation
        End Get
    End Property

    Protected Sub New(ByVal type As TreeType, ByVal attributes As AttributeBlockCollection, ByVal modifiers As ModifierCollection, ByVal keywordLocation As Location, ByVal subOrFunctionLocation As Location, ByVal name As SimpleName, ByVal typeParameters As TypeParameterCollection, ByVal parameters As ParameterCollection, ByVal asLocation As Location, ByVal resultTypeAttributes As AttributeBlockCollection, ByVal resultType As TypeName, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(type, attributes, modifiers, keywordLocation, name, typeParameters, parameters, asLocation, resultTypeAttributes, resultType, span, comments)

        Debug.Assert(type = TreeType.DelegateSubDeclaration OrElse type = TreeType.DelegateFunctionDeclaration)

        _SubOrFunctionLocation = subOrFunctionLocation
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)
    End Sub
End Class