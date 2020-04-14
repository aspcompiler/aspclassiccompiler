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
''' A parse tree for an AddHandler or RemoveHandler statement.
''' </summary>
Public MustInherit Class HandlerStatement
    Inherits Statement

    Private ReadOnly _Name As Expression
    Private ReadOnly _CommaLocation As Location
    Private ReadOnly _DelegateExpression As Expression

    ''' <summary>
    ''' The event name.
    ''' </summary>
    Public ReadOnly Property Name() As Expression
        Get
            Return _Name
        End Get
    End Property

    ''' <summary>
    ''' The location of the ','.
    ''' </summary>
    Public ReadOnly Property CommaLocation() As Location
        Get
            Return _CommaLocation
        End Get
    End Property

    ''' <summary>
    ''' The delegate expression.
    ''' </summary>
    Public ReadOnly Property DelegateExpression() As Expression
        Get
            Return _DelegateExpression
        End Get
    End Property

    Protected Sub New(ByVal type As TreeType, ByVal name As Expression, ByVal commaLocation As Location, ByVal delegateExpression As Expression, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(type, span, comments)

        Debug.Assert(type = TreeType.AddHandlerStatement OrElse type = TreeType.RemoveHandlerStatement)

        If name Is Nothing Then
            Throw New ArgumentNullException("name")
        End If

        If delegateExpression Is Nothing Then
            Throw New ArgumentNullException("delegateExpression")
        End If

        SetParent(name)
        SetParent(delegateExpression)

        _Name = name
        _CommaLocation = commaLocation
        _DelegateExpression = delegateExpression
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Name)
        AddChild(childList, DelegateExpression)
    End Sub
End Class