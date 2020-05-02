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
''' A parse tree for an attribute usage.
''' </summary>
Public NotInheritable Class Attribute
    Inherits Tree

    Private ReadOnly _AttributeType As AttributeTypes
    Private ReadOnly _AttributeTypeLocation As Location
    Private ReadOnly _ColonLocation As Location
    Private ReadOnly _Name As Name
    Private ReadOnly _Arguments As ArgumentCollection

    ''' <summary>
    ''' The target type of the attribute.
    ''' </summary>
    Public ReadOnly Property AttributeType() As AttributeTypes
        Get
            Return _AttributeType
        End Get
    End Property

    ''' <summary>
    ''' The location of the attribute type, if any.
    ''' </summary>
    Public ReadOnly Property AttributeTypeLocation() As Location
        Get
            Return _AttributeTypeLocation
        End Get
    End Property

    ''' <summary>
    ''' The location of the ':', if any.
    ''' </summary>
    Public ReadOnly Property ColonLocation() As Location
        Get
            Return _ColonLocation
        End Get
    End Property

    ''' <summary>
    ''' The name of the attribute being applied.
    ''' </summary>
    Public ReadOnly Property Name() As Name
        Get
            Return _Name
        End Get
    End Property

    ''' <summary>
    ''' The arguments to the attribute.
    ''' </summary>
    Public ReadOnly Property Arguments() As ArgumentCollection
        Get
            Return _Arguments
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new attribute parse tree.
    ''' </summary>
    ''' <param name="attributeType">The target type of the attribute.</param>
    ''' <param name="attributeTypeLocation">The location of the attribute type.</param>
    ''' <param name="colonLocation">The location of the ':'.</param>
    ''' <param name="name">The name of the attribute being applied.</param>
    ''' <param name="arguments">The arguments to the attribute.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal attributeType As AttributeTypes, ByVal attributeTypeLocation As Location, ByVal colonLocation As Location, ByVal name As Name, ByVal arguments As ArgumentCollection, ByVal span As Span)
        MyBase.New(TreeType.Attribute, span)

        If name Is Nothing Then
            Throw New ArgumentNullException("name")
        End If

        SetParent(name)
        SetParent(arguments)

        _AttributeType = attributeType
        _AttributeTypeLocation = attributeTypeLocation
        _ColonLocation = colonLocation
        _Name = name
        _Arguments = arguments
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Name)
        AddChild(childList, Arguments)
    End Sub
End Class
