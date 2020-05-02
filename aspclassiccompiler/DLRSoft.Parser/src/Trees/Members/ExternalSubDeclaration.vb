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
''' A parse tree for a Declare Sub statement.
''' </summary>
Public NotInheritable Class ExternalSubDeclaration
    Inherits ExternalDeclaration

    ''' <summary>
    ''' Constructs a parse tree for a Declare Sub statement.
    ''' </summary>
    ''' <param name="attributes">The attributes for the parse tree.</param>
    ''' <param name="modifiers">The modifiers for the parse tree.</param>
    ''' <param name="keywordLocation">The location of the keyword.</param>
    ''' <param name="charsetLocation">The location of the 'Ansi', 'Auto' or 'Unicode', if any.</param>
    ''' <param name="charset">The charset.</param>
    ''' <param name="subLocation">The location of 'Sub'.</param>
    ''' <param name="name">The name of the declaration.</param>
    ''' <param name="libLocation">The location of 'Lib', if any.</param>
    ''' <param name="libLiteral">The library, if any.</param>
    ''' <param name="aliasLocation">The location of 'Alias', if any.</param>
    ''' <param name="aliasLiteral">The alias, if any.</param>
    ''' <param name="parameters">The parameters of the declaration.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal attributes As AttributeBlockCollection, ByVal modifiers As ModifierCollection, ByVal keywordLocation As Location, ByVal charsetLocation As Location, ByVal charset As Charset, ByVal subLocation As Location, ByVal name As SimpleName, ByVal libLocation As Location, ByVal libLiteral As StringLiteralExpression, ByVal aliasLocation As Location, ByVal aliasLiteral As StringLiteralExpression, ByVal parameters As ParameterCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.ExternalSubDeclaration, attributes, modifiers, keywordLocation, charsetLocation, charset, subLocation, name, libLocation, libLiteral, aliasLocation, aliasLiteral, parameters, Nothing, Nothing, Nothing, span, comments)
    End Sub
End Class