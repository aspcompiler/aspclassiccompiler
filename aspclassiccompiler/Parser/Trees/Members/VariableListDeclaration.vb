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
''' A parse tree for variable declarations.
''' </summary>
Public NotInheritable Class VariableListDeclaration
    Inherits ModifiedDeclaration

    Private ReadOnly _VariableDeclarators As VariableDeclaratorCollection

    ''' <summary>
    ''' The variables being declared.
    ''' </summary>
    Public ReadOnly Property VariableDeclarators() As VariableDeclaratorCollection
        Get
            Return _VariableDeclarators
        End Get
    End Property

    ''' <summary>
    ''' Constructs a parse tree for variable declarations.
    ''' </summary>
    ''' <param name="attributes">The attributes on the declaration.</param>
    ''' <param name="modifiers">The modifiers on the declaration.</param>
    ''' <param name="variableDeclarators">The variables being declared.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal attributes As AttributeBlockCollection, ByVal modifiers As ModifierCollection, ByVal variableDeclarators As VariableDeclaratorCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.VariableListDeclaration, attributes, modifiers, span, comments)

        If variableDeclarators Is Nothing Then
            Throw New ArgumentNullException("variableDeclarators")
        End If

        SetParent(variableDeclarators)

        _VariableDeclarators = variableDeclarators
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)
        AddChild(childList, VariableDeclarators)
    End Sub
End Class