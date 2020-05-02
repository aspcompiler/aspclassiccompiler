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
''' A parse tree for a type parameter.
''' </summary>
Public NotInheritable Class TypeParameter
    Inherits Tree

    Private ReadOnly _TypeName As SimpleName
    Private ReadOnly _AsLocation As Location
    Private ReadOnly _TypeConstraints As TypeConstraintCollection

    ''' <summary>
    ''' The name of the type parameter.
    ''' </summary>
    Public ReadOnly Property TypeName() As SimpleName
        Get
            Return _TypeName
        End Get
    End Property

    ''' <summary>
    ''' The location of the 'As', if any.
    ''' </summary>
    Public ReadOnly Property AsLocation() As Location
        Get
            Return _AsLocation
        End Get
    End Property

    ''' <summary>
    ''' The constraints, if any.
    ''' </summary>
    Public ReadOnly Property TypeConstraints() As TypeConstraintCollection
        Get
            Return _TypeConstraints
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parameter parse tree.
    ''' </summary>
    ''' <param name="typeName">The name of the type parameter.</param>
    ''' <param name="asLocation">The location of the 'As'.</param>
    ''' <param name="typeConstraints">The constraints on the type parameter. Can be Nothing.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal typeName As SimpleName, ByVal asLocation As Location, ByVal typeConstraints As TypeConstraintCollection, ByVal span As Span)
        MyBase.New(TreeType.TypeParameter, span)

        If typeName Is Nothing Then
            Throw New ArgumentNullException("typeName")
        End If

        SetParent(typeName)
        SetParent(typeConstraints)

        _TypeName = typeName
        _AsLocation = asLocation
        _TypeConstraints = typeConstraints
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, TypeName)
        AddChild(childList, TypeConstraints)
    End Sub
End Class
