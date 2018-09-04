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
''' A parse tree for a delegate Sub declaration.
''' </summary>
Public NotInheritable Class DelegateSubDeclaration
    Inherits DelegateDeclaration

    ''' <summary>
    ''' Constructs a new parse tree for a delegate Sub declaration.
    ''' </summary>
    ''' <param name="attributes">The attributes for the parse tree.</param>
    ''' <param name="modifiers">The modifiers for the parse tree.</param>
    ''' <param name="keywordLocation">The location of the keyword.</param>
    ''' <param name="subLocation">The location of the 'Sub'.</param>
    ''' <param name="name">The name of the declaration.</param>
    ''' <param name="typeParameters">The type parameters of the declaration, if any.</param>
    ''' <param name="parameters">The parameters of the declaration.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal attributes As AttributeBlockCollection, ByVal modifiers As ModifierCollection, ByVal keywordLocation As Location, ByVal subLocation As Location, ByVal name As SimpleName, ByVal typeParameters As TypeParameterCollection, ByVal parameters As ParameterCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.DelegateSubDeclaration, attributes, modifiers, keywordLocation, subLocation, name, typeParameters, parameters, Nothing, Nothing, Nothing, span, comments)
    End Sub
End Class