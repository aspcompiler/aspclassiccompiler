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
''' A parse tree for a variable declarator (e.g. "x As Integer")
''' </summary>
Public NotInheritable Class VariableDeclarator
    Inherits Tree

    Private ReadOnly _VariableNames As VariableNameCollection
    Private ReadOnly _AsLocation As Location
    Private ReadOnly _NewLocation As Location
    Private ReadOnly _VariableType As TypeName
    Private ReadOnly _Arguments As ArgumentCollection
    Private ReadOnly _EqualsLocation As Location
    Private ReadOnly _Initializer As Initializer

    ''' <summary>
    ''' The variable names being declared.
    ''' </summary>
    Public ReadOnly Property VariableNames() As VariableNameCollection
        Get
            Return _VariableNames
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
    ''' The location of the 'New', if any.
    ''' </summary>
    Public ReadOnly Property NewLocation() As Location
        Get
            Return _NewLocation
        End Get
    End Property

    ''' <summary>
    ''' The type of the variables being declared, if any.
    ''' </summary>
    Public ReadOnly Property VariableType() As TypeName
        Get
            Return _VariableType
        End Get
    End Property

    ''' <summary>
    ''' The arguments to the constructor, if any.
    ''' </summary>
    Public ReadOnly Property Arguments() As ArgumentCollection
        Get
            Return _Arguments
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
    ''' The variable initializer, if any.
    ''' </summary>
    Public ReadOnly Property Initializer() As Initializer
        Get
            Return _Initializer
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a variable declarator.
    ''' </summary>
    ''' <param name="variableNames">The names of the variables being declared.</param>
    ''' <param name="asLocation">The location of the 'As', if any.</param>
    ''' <param name="newLocation">The location of the 'New', if any.</param>
    ''' <param name="variableType">The type of the variables being declared, if any.</param>
    ''' <param name="arguments">The arguments of the constructor, if any.</param>
    ''' <param name="equalsLocation">The location of the '=', if any.</param>
    ''' <param name="initializer">The variable initializer, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal variableNames As VariableNameCollection, ByVal asLocation As Location, ByVal newLocation As Location, ByVal variableType As TypeName, ByVal arguments As ArgumentCollection, ByVal equalsLocation As Location, ByVal initializer As Initializer, ByVal span As Span)
        MyBase.New(TreeType.VariableDeclarator, span)

        If variableNames Is Nothing Then
            Throw New ArgumentNullException("variableNames")
        End If

        SetParent(variableNames)
        SetParent(variableType)
        SetParent(arguments)
        SetParent(initializer)

        _VariableNames = variableNames
        _AsLocation = asLocation
        _NewLocation = newLocation
        _VariableType = variableType
        _Arguments = arguments
        _EqualsLocation = equalsLocation
        _Initializer = initializer
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, VariableNames)
        AddChild(childList, VariableType)
        AddChild(childList, Arguments)
        AddChild(childList, Initializer)
    End Sub
End Class