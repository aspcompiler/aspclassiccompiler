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
''' A parse tree for an argument to a call or index.
''' </summary>
Public NotInheritable Class Argument
    Inherits Tree

    Private ReadOnly _Name As SimpleName
    Private ReadOnly _ColonEqualsLocation As Location
    Private ReadOnly _Expression As Expression

    ''' <summary>
    ''' The name of the argument, if any.
    ''' </summary>
    Public ReadOnly Property Name() As SimpleName
        Get
            Return _Name
        End Get
    End Property

    ''' <summary>
    ''' The location of the ':=', if any.
    ''' </summary>
    Public ReadOnly Property ColonEqualsLocation() As Location
        Get
            Return _ColonEqualsLocation
        End Get
    End Property

    ''' <summary>
    ''' The argument, if any.
    ''' </summary>
    Public ReadOnly Property Expression() As Expression
        Get
            Return _Expression
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for an argument.
    ''' </summary>
    ''' <param name="name">The name of the argument, if any.</param>
    ''' <param name="colonEqualsLocation">The location of the ':=', if any.</param>
    ''' <param name="expression">The expression, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal name As SimpleName, ByVal colonEqualsLocation As Location, ByVal expression As Expression, ByVal span As Span)
        MyBase.New(TreeType.Argument, span)

        If expression Is Nothing Then
            Throw New ArgumentNullException("expression")
        End If

        SetParent(name)
        SetParent(expression)

        _Name = name
        _ColonEqualsLocation = colonEqualsLocation
        _Expression = expression
    End Sub

    Private Sub New()
        MyBase.New(TreeType.Argument, Nothing)
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Name)
        AddChild(childList, Expression)
    End Sub
End Class