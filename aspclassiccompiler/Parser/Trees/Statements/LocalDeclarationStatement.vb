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
''' A parse tree for a local declaration statement.
''' </summary>
Public NotInheritable Class LocalDeclarationStatement
    Inherits Statement

    Private ReadOnly _Modifiers As ModifierCollection
    Private ReadOnly _VariableDeclarators As VariableDeclaratorCollection

    ''' <summary>
    ''' The statement modifiers.
    ''' </summary>
    Public ReadOnly Property Modifiers() As ModifierCollection
        Get
            Return _Modifiers
        End Get
    End Property

    ''' <summary>
    ''' The variable declarators.
    ''' </summary>
    Public ReadOnly Property VariableDeclarators() As VariableDeclaratorCollection
        Get
            Return _VariableDeclarators
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a local declaration statement.
    ''' </summary>
    ''' <param name="modifiers">The statement modifiers.</param>
    ''' <param name="variableDeclarators">The variable declarators.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal modifiers As ModifierCollection, ByVal variableDeclarators As VariableDeclaratorCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.LocalDeclarationStatement, span, comments)

        If modifiers Is Nothing Then
            Throw New ArgumentNullException("modifers")
        End If

        If variableDeclarators Is Nothing Then
            Throw New ArgumentNullException("variableDeclarators")
        End If

        SetParent(modifiers)
        SetParent(variableDeclarators)

        _Modifiers = modifiers
        _VariableDeclarators = variableDeclarators
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Modifiers)
        AddChild(childList, VariableDeclarators)
    End Sub
End Class