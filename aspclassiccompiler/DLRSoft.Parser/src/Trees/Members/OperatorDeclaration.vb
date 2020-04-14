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
''' A parse tree for an overloaded operator declaration.
''' </summary>
Public NotInheritable Class OperatorDeclaration
    Inherits MethodDeclaration

    Private ReadOnly _OperatorToken As Token

    ''' <summary>
    ''' The operator being overloaded.
    ''' </summary>
    Public ReadOnly Property OperatorToken() As Token
        Get
            Return _OperatorToken
        End Get
    End Property

    ''' <summary>
    ''' Creates a new parse tree for an overloaded operator declaration.
    ''' </summary>
    ''' <param name="attributes">The attributes for the parse tree.</param>
    ''' <param name="modifiers">The modifiers for the parse tree.</param>
    ''' <param name="keywordLocation">The location of the keyword.</param>
    ''' <param name="operatorToken">The operator being overloaded.</param>
    ''' <param name="parameters">The parameters of the declaration.</param>
    ''' <param name="asLocation">The location of the 'As', if any.</param>
    ''' <param name="resultTypeAttributes">The attributes on the result type, if any.</param>
    ''' <param name="resultType">The result type, if any.</param>
    ''' <param name="statements">The statements in the declaration.</param>
    ''' <param name="endDeclaration">The end block declaration, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal attributes As AttributeBlockCollection, ByVal modifiers As ModifierCollection, ByVal keywordLocation As Location, ByVal operatorToken As Token, ByVal parameters As ParameterCollection, ByVal asLocation As Location, ByVal resultTypeAttributes As AttributeBlockCollection, ByVal resultType As TypeName, ByVal statements As StatementCollection, ByVal endDeclaration As EndBlockDeclaration, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.OperatorDeclaration, attributes, modifiers, keywordLocation, Nothing, Nothing, parameters, asLocation, resultTypeAttributes, resultType, Nothing, Nothing, statements, endDeclaration, span, comments)

        _OperatorToken = operatorToken
    End Sub
End Class
