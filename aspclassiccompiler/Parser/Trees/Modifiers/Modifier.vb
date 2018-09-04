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
''' A parse tree for a declaration modifier.
''' </summary>
Public NotInheritable Class Modifier
    Inherits Tree

    Private ReadOnly _ModifierType As ModifierTypes

    ''' <summary>
    ''' The type of the modifier.
    ''' </summary>
    Public ReadOnly Property ModifierType() As ModifierTypes
        Get
            Return _ModifierType
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new modifier parse tree.
    ''' </summary>
    ''' <param name="modifierType">The type of the modifier.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal modifierType As ModifierTypes, ByVal span As Span)
        MyBase.New(TreeType.Modifier, span)

        If (modifierType And (modifierType - 1)) <> 0 OrElse _
           modifierType < ModifierTypes.None OrElse _
           modifierType > ModifierTypes.Narrowing Then
            Throw New ArgumentOutOfRangeException("modifierType")
        End If

        _ModifierType = modifierType
    End Sub
End Class