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
''' A parse tree for a parameter.
''' </summary>
Public NotInheritable Class Parameter
    Inherits Tree

    Private ReadOnly _Attributes As AttributeBlockCollection
    Private ReadOnly _Modifiers As ModifierCollection
    Private ReadOnly _VariableName As VariableName
    Private ReadOnly _AsLocation As Location
    Private ReadOnly _ParameterType As TypeName
    Private ReadOnly _EqualsLocation As Location
    Private ReadOnly _Initializer As Initializer

    ''' <summary>
    ''' The attributes on the parameter.
    ''' </summary>
    Public ReadOnly Property Attributes() As AttributeBlockCollection
        Get
            Return _Attributes
        End Get
    End Property

    ''' <summary>
    ''' The modifiers on the parameter.
    ''' </summary>
    Public ReadOnly Property Modifiers() As ModifierCollection
        Get
            Return _Modifiers
        End Get
    End Property

    ''' <summary>
    ''' The name of the parameter.
    ''' </summary>
    Public ReadOnly Property VariableName() As VariableName
        Get
            Return _VariableName
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
    ''' The parameter type, if any.
    ''' </summary>
    Public ReadOnly Property ParameterType() As TypeName
        Get
            Return _ParameterType
        End Get
    End Property

    ''' <summary>
    ''' The location of the '=', if any.
    ''' </summary>
    Public ReadOnly Property EqualsLocation() As Location
        Get
            Return _EqualsLocation
        End Get
    End Property

    ''' <summary>
    ''' The initializer for the parameter, if any.
    ''' </summary>
    Public ReadOnly Property Initializer() As Initializer
        Get
            Return _Initializer
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parameter parse tree.
    ''' </summary>
    ''' <param name="attributes">The attributes on the parameter.</param>
    ''' <param name="modifiers">The modifiers on the parameter.</param>
    ''' <param name="variableName">The name of the parameter.</param>
    ''' <param name="asLocation">The location of the 'As'.</param>
    ''' <param name="parameterType">The type of the parameter. Can be Nothing.</param>
    ''' <param name="equalsLocation">The location of the '='.</param>
    ''' <param name="initializer">The initializer for the parameter. Can be Nothing.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal attributes As AttributeBlockCollection, ByVal modifiers As ModifierCollection, ByVal variableName As VariableName, ByVal asLocation As Location, ByVal parameterType As TypeName, ByVal equalsLocation As Location, ByVal initializer As Initializer, ByVal span As Span)
        MyBase.New(TreeType.Parameter, span)

        If variableName Is Nothing Then
            Throw New ArgumentNullException("variableName")
        End If

        SetParent(attributes)
        SetParent(modifiers)
        SetParent(variableName)
        SetParent(parameterType)
        SetParent(initializer)

        _Attributes = attributes
        _Modifiers = modifiers
        _VariableName = variableName
        _AsLocation = asLocation
        _ParameterType = parameterType
        _EqualsLocation = equalsLocation
        _Initializer = initializer
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Attributes)
        AddChild(childList, Modifiers)
        AddChild(childList, VariableName)
        AddChild(childList, ParameterType)
        AddChild(childList, Initializer)
    End Sub
End Class