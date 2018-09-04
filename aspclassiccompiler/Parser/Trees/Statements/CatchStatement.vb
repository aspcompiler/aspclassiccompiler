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
''' A parse tree for a Catch statement.
''' </summary>
Public NotInheritable Class CatchStatement
    Inherits Statement

    Private ReadOnly _Name As SimpleName
    Private ReadOnly _AsLocation As Location
    Private ReadOnly _ExceptionType As TypeName
    Private ReadOnly _WhenLocation As Location
    Private ReadOnly _FilterExpression As Expression

    ''' <summary>
    ''' The name of the catch variable, if any.
    ''' </summary>
    Public ReadOnly Property Name() As SimpleName
        Get
            Return _Name
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
    ''' The type of the catch variable, if any.
    ''' </summary>
    Public ReadOnly Property ExceptionType() As TypeName
        Get
            Return _ExceptionType
        End Get
    End Property

    ''' <summary>
    ''' The location of the 'When', if any.
    ''' </summary>
    Public ReadOnly Property WhenLocation() As Location
        Get
            Return _WhenLocation
        End Get
    End Property

    ''' <summary>
    ''' The filter expression, if any.
    ''' </summary>
    Public ReadOnly Property FilterExpression() As Expression
        Get
            Return _FilterExpression
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a Catch statement.
    ''' </summary>
    ''' <param name="name">The name of the catch variable, if any.</param>
    ''' <param name="asLocation">The location of the 'As', if any.</param>
    ''' <param name="exceptionType">The type of the catch variable, if any.</param>
    ''' <param name="whenLocation">The location of the 'When', if any.</param>
    ''' <param name="filterExpression">The filter expression, if any.</param>
    ''' <param name="span">The location of the parse tree, if any.</param>
    ''' <param name="comments">The comments for the parse tree, if any.</param>
    Public Sub New(ByVal name As SimpleName, ByVal asLocation As Location, ByVal exceptionType As TypeName, ByVal whenLocation As Location, ByVal filterExpression As Expression, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.CatchStatement, span, comments)

        SetParent(name)
        SetParent(exceptionType)
        SetParent(filterExpression)

        _Name = name
        _AsLocation = asLocation
        _ExceptionType = exceptionType
        _WhenLocation = whenLocation
        _FilterExpression = filterExpression
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Name)
        AddChild(childList, ExceptionType)
        AddChild(childList, FilterExpression)
    End Sub
End Class