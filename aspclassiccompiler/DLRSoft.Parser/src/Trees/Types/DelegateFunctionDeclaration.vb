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
''' A parse tree for a delegate Function declaration.
''' </summary>
Public NotInheritable Class DelegateFunctionDeclaration
    Inherits DelegateDeclaration

    ''' <summary>
    ''' Constructs a new parse tree for a delegate declaration.
    ''' </summary>
    ''' <param name="attributes">The attributes for the parse tree.</param>
    ''' <param name="modifiers">The modifiers for the parse tree.</param>
    ''' <param name="keywordLocation">The location of the keyword.</param>
    ''' <param name="functionLocation">The location of the 'Function'.</param>
    ''' <param name="name">The name of the declaration.</param>
    ''' <param name="typeParameters">The type parameters of the declaration, if any.</param>
    ''' <param name="parameters">The parameters of the declaration.</param>
    ''' <param name="asLocation">The location of the 'As', if any.</param>
    ''' <param name="resultTypeAttributes">The attributes on the result type, if any.</param>
    ''' <param name="resultType">The result type, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal attributes As AttributeBlockCollection, ByVal modifiers As ModifierCollection, ByVal keywordLocation As Location, ByVal functionLocation As Location, ByVal name As SimpleName, ByVal typeParameters As TypeParameterCollection, ByVal parameters As ParameterCollection, ByVal asLocation As Location, ByVal resultTypeAttributes As AttributeBlockCollection, ByVal resultType As TypeName, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.DelegateFunctionDeclaration, attributes, modifiers, keywordLocation, functionLocation, name, typeParameters, parameters, asLocation, resultTypeAttributes, resultType, span, comments)
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)
    End Sub
End Class