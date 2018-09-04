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
''' A parse tree for a statement.
''' </summary>
Public MustInherit Class Statement
    Inherits Tree

    Private ReadOnly _Comments As ReadOnlyCollection(Of Comment)

    ''' <summary>
    ''' The comments for the tree.
    ''' </summary>
    Public ReadOnly Property Comments() As ReadOnlyCollection(Of Comment)
        Get
            Return _Comments
        End Get
    End Property

    Protected Sub New(ByVal type As TreeType, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(type, span)

        'LC Allow declarations to be craeted as statement
        Debug.Assert((type >= TreeType.EmptyStatement AndAlso type <= TreeType.EndBlockStatement) OrElse (type >= TreeType.EmptyDeclaration AndAlso type <= TreeType.DelegateFunctionDeclaration))

        If comments IsNot Nothing Then
            _Comments = New ReadOnlyCollection(Of Comment)(comments)
        End If
    End Sub
End Class