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
''' A read-only collection of modifiers.
''' </summary>
Public NotInheritable Class ModifierCollection
    Inherits TreeCollection(Of Modifier)

    Private ReadOnly _ModifierTypes As ModifierTypes

    ''' <summary>
    ''' All the modifiers in the collection.
    ''' </summary>
    Public ReadOnly Property ModifierTypes() As ModifierTypes
        Get
            Return _ModifierTypes
        End Get
    End Property

    ''' <summary>
    ''' Constructs a collection of modifiers.
    ''' </summary>
    ''' <param name="modifiers">The modifiers in the collection.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal modifiers As IList(Of Modifier), ByVal span As Span)
        MyBase.New(TreeType.ModifierCollection, modifiers, span)

        If modifiers Is Nothing OrElse modifiers.Count = 0 Then
            Throw New ArgumentException("ModifierCollection cannot be empty.")
        End If

        For Each Modifier As Modifier In modifiers
            _ModifierTypes = _ModifierTypes Or Modifier.ModifierType
        Next
    End Sub
End Class