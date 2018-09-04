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
''' A parse tree for a Function declaration.
''' </summary>
Public NotInheritable Class FunctionDeclaration
    Inherits MethodDeclaration

    ''' <summary>
    ''' Creates a new parse tree for a Function declaration.
    ''' </summary>
    ''' <param name="attributes">The attributes for the parse tree.</param>
    ''' <param name="modifiers">The modifiers for the parse tree.</param>
    ''' <param name="keywordLocation">The location of the keyword.</param>
    ''' <param name="name">The name of the declaration.</param>
    ''' <param name="typeParameters">The type parameters on the declaration, if any.</param>
    ''' <param name="parameters">The parameters of the declaration.</param>
    ''' <param name="asLocation">The location of the 'As', if any.</param>
    ''' <param name="resultTypeAttributes">The attributes on the result type, if any.</param>
    ''' <param name="resultType">The result type, if any.</param>
    ''' <param name="implementsList">The list of implemented members.</param>
    ''' <param name="handlesList">The list of handled events.</param>
    ''' <param name="statements">The statements in the declaration.</param>
    ''' <param name="endDeclaration">The end block declaration, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal attributes As AttributeBlockCollection, ByVal modifiers As ModifierCollection, ByVal keywordLocation As Location, ByVal name As SimpleName, ByVal typeParameters As TypeParameterCollection, ByVal parameters As ParameterCollection, ByVal asLocation As Location, ByVal resultTypeAttributes As AttributeBlockCollection, ByVal resultType As TypeName, ByVal implementsList As NameCollection, ByVal handlesList As NameCollection, ByVal statements As StatementCollection, ByVal endDeclaration As EndBlockDeclaration, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.FunctionDeclaration, attributes, modifiers, keywordLocation, name, typeParameters, parameters, asLocation, resultTypeAttributes, resultType, implementsList, handlesList, statements, endDeclaration, span, comments)
    End Sub
End Class
