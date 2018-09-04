'
' Visual Basic .NET Parser
'
' Copyright (C) 2005, Microsoft Corporation. All rights reserved.
'
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
' MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
'
'LC Changed Declaration to Inherits from Statement instead of tree so that we can add declaration to statement collection

''' <summary>
''' A parse tree for a declaration.
''' </summary>
Public Class Declaration
    Inherits Statement

    ''' <summary>
    ''' Creates a bad declaration.
    ''' </summary>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    ''' <returns>A bad declaration.</returns>
    Public Shared Function GetBadDeclaration(ByVal span As Span, ByVal comments As IList(Of Comment)) As Declaration
        Return New Declaration(span, comments)
    End Function

    Protected Sub New(ByVal type As TreeType, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(type, span, comments)

        Debug.Assert(type >= TreeType.EmptyDeclaration AndAlso type <= TreeType.DelegateFunctionDeclaration)

    End Sub

    Private Sub New(ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.SyntaxError, span, comments)
    End Sub

    Public Overrides ReadOnly Property IsBad() As Boolean
        Get
            Return Type = TreeType.SyntaxError
        End Get
    End Property
End Class