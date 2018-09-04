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
''' A parse tree for an assembly-level or module-level attribute declaration.
''' </summary>
Public NotInheritable Class AttributeDeclaration
    Inherits Declaration

    Private ReadOnly _Attributes As AttributeBlockCollection

    ''' <summary>
    ''' The attributes.
    ''' </summary>
    Public ReadOnly Property Attributes() As AttributeBlockCollection
        Get
            Return _Attributes
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for assembly-level or module-level attribute declarations.
    ''' </summary>
    ''' <param name="attributes">The attributes.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal attributes As AttributeBlockCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.AttributeDeclaration, span, comments)

        If attributes Is Nothing Then
            Throw New ArgumentNullException("attributes")
        End If

        SetParent(attributes)

        _Attributes = attributes
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Attributes)
    End Sub
End Class