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
''' A parse tree for an Imports declaration.
''' </summary>
Public NotInheritable Class ImportsDeclaration
    Inherits Declaration

    Private ReadOnly _ImportMembers As ImportCollection

    ''' <summary>
    ''' The members imported.
    ''' </summary>
    Public ReadOnly Property ImportMembers() As ImportCollection
        Get
            Return _ImportMembers
        End Get
    End Property

    ''' <summary>
    ''' Constructs a parse tree for an Imports declaration.
    ''' </summary>
    ''' <param name="importMembers">The members imported.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal importMembers As ImportCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.ImportsDeclaration, span, comments)

        If importMembers Is Nothing Then
            Throw New ArgumentNullException("importMembers")
        End If

        SetParent(importMembers)

        _ImportMembers = importMembers
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)

        AddChild(childList, ImportMembers)
    End Sub
End Class