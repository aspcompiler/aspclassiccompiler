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
''' A parse tree for an Implements declaration.
''' </summary>
Public NotInheritable Class ImplementsDeclaration
    Inherits Declaration

    Private ReadOnly _ImplementedTypes As TypeNameCollection

    ''' <summary>
    ''' The list of types.
    ''' </summary>
    Public ReadOnly Property ImplementedTypes() As TypeNameCollection
        Get
            Return _ImplementedTypes
        End Get
    End Property

    ''' <summary>
    ''' Constructs a parse tree for an Implements declaration.
    ''' </summary>
    ''' <param name="implementedTypes">The types inherited or implemented.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal implementedTypes As TypeNameCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.ImplementsDeclaration, span, comments)

        If implementedTypes Is Nothing Then
            Throw New ArgumentNullException("implementedTypes")
        End If

        SetParent(implementedTypes)

        _ImplementedTypes = implementedTypes
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)

        AddChild(childList, ImplementedTypes)
    End Sub
End Class