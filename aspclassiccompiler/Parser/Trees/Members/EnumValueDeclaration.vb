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
''' A parse tree for an enumerated value declaration.
''' </summary>
Public NotInheritable Class EnumValueDeclaration
    Inherits ModifiedDeclaration

    Private ReadOnly _Name As Name
    Private ReadOnly _EqualsLocation As Location
    Private ReadOnly _Expression As Expression

    ''' <summary>
    ''' The name of the enumerated value.
    ''' </summary>
    Public ReadOnly Property Name() As Name
        Get
            Return _Name
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
    ''' The enumerated value, if any.
    ''' </summary>
    Public ReadOnly Property Expression() As Expression
        Get
            Return _Expression
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for an enumerated value.
    ''' </summary>
    ''' <param name="attributes">The attributes on the declaration.</param>
    ''' <param name="name">The name of the declaration.</param>
    ''' <param name="equalsLocation">The location of the '=', if any.</param>
    ''' <param name="expression">The enumerated value, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal attributes As AttributeBlockCollection, ByVal name As Name, ByVal equalsLocation As Location, ByVal expression As Expression, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.EnumValueDeclaration, attributes, Nothing, span, comments)

        If name Is Nothing Then
            Throw New ArgumentNullException("name")
        End If

        SetParent(name)
        SetParent(expression)

        _Name = name
        _EqualsLocation = equalsLocation
        _Expression = expression
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)

        AddChild(childList, Name)
        AddChild(childList, Expression)
    End Sub
End Class