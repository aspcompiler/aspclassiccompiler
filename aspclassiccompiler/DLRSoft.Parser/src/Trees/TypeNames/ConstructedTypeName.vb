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
''' A parse tree for a constructed generic type name.
''' </summary>
Public NotInheritable Class ConstructedTypeName
    Inherits NamedTypeName

    Private ReadOnly _TypeArguments As TypeArgumentCollection

    ''' <summary>
    ''' The type arguments.
    ''' </summary>
    Public ReadOnly Property TypeArguments() As TypeArgumentCollection
        Get
            Return _TypeArguments
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a generic constructed type name.
    ''' </summary>
    ''' <param name="name">The generic type being constructed.</param>
    ''' <param name="typeArguments">The type arguments.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal name As Name, ByVal typeArguments As TypeArgumentCollection, ByVal span As Span)
        MyBase.New(TreeType.ConstructedType, name, span)

        If typeArguments Is Nothing Then
            Throw New ArgumentNullException("typeArguments")
        End If

        SetParent(typeArguments)
        _TypeArguments = typeArguments
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, TypeArguments)
    End Sub
End Class