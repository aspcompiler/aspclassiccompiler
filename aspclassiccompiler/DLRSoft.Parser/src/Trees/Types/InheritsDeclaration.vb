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
''' A parse tree for an Inherits declaration.
''' </summary>
Public NotInheritable Class InheritsDeclaration
    Inherits Declaration

    Private ReadOnly _InheritedTypes As TypeNameCollection

    ''' <summary>
    ''' The list of types.
    ''' </summary>
    Public ReadOnly Property InheritedTypes() As TypeNameCollection
        Get
            Return _InheritedTypes
        End Get
    End Property

    ''' <summary>
    ''' Constructs a parse tree for an Inherits declaration.
    ''' </summary>
    ''' <param name="inheritedTypes">The types inherited or implemented.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal inheritedTypes As TypeNameCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.InheritsDeclaration, span, comments)

        If inheritedTypes Is Nothing Then
            Throw New ArgumentNullException("inheritedTypes")
        End If

        SetParent(inheritedTypes)

        _InheritedTypes = inheritedTypes
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)

        AddChild(childList, InheritedTypes)
    End Sub
End Class