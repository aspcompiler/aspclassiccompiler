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
''' A parse tree for an Else If statement.
''' </summary>
Public NotInheritable Class ElseIfStatement
    Inherits Statement

    Private ReadOnly _Expression As Expression
    Private ReadOnly _ThenLocation As Location

    ''' <summary>
    ''' The conditional expression.
    ''' </summary>
    Public ReadOnly Property Expression() As Expression
        Get
            Return _Expression
        End Get
    End Property

    ''' <summary>
    ''' The location of the 'Then', if any.
    ''' </summary>
    Public ReadOnly Property ThenLocation() As Location
        Get
            Return _ThenLocation
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for an Else If statement.
    ''' </summary>
    ''' <param name="expression">The conditional expression.</param>
    ''' <param name="thenLocation">The location of the 'Then', if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal expression As Expression, ByVal thenLocation As Location, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.ElseIfStatement, span, comments)

        If expression Is Nothing Then
            Throw New ArgumentNullException("expression")
        End If

        SetParent(expression)
        _Expression = expression
        _ThenLocation = thenLocation
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Expression)
    End Sub
End Class