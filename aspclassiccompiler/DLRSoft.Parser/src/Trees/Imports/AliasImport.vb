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
''' A parse tree for an Imports statement that aliases a type or namespace.
''' </summary>
Public NotInheritable Class AliasImport
    Inherits Import

    Private ReadOnly _Name As SimpleName
    Private ReadOnly _EqualsLocation As Location
    Private ReadOnly _AliasedTypeName As TypeName

    ''' <summary>
    ''' The alias name.
    ''' </summary>
    Public ReadOnly Property Name() As SimpleName
        Get
            Return _Name
        End Get
    End Property

    ''' <summary>
    ''' The location of the '='.
    ''' </summary>
    Public ReadOnly Property EqualsLocation() As Location
        Get
            Return _EqualsLocation
        End Get
    End Property

    ''' <summary>
    ''' The name being aliased.
    ''' </summary>
    Public ReadOnly Property AliasedTypeName() As TypeName
        Get
            Return _AliasedTypeName
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new aliased import parse tree.
    ''' </summary>
    ''' <param name="name">The name of the alias.</param>
    ''' <param name="equalsLocation">The location of the '='.</param>
    ''' <param name="aliasedTypeName">The name being aliased.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal name As SimpleName, ByVal equalsLocation As Location, ByVal aliasedTypeName As TypeName, ByVal span As Span)
        MyBase.New(TreeType.AliasImport, span)

        If aliasedTypeName Is Nothing Then
            Throw New ArgumentNullException("aliasedTypeName")
        End If

        If name Is Nothing Then
            Throw New ArgumentNullException("name")
        End If

        SetParent(name)
        SetParent(aliasedTypeName)

        _Name = name
        _EqualsLocation = equalsLocation
        _AliasedTypeName = aliasedTypeName
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)

        AddChild(childList, AliasedTypeName)
    End Sub
End Class