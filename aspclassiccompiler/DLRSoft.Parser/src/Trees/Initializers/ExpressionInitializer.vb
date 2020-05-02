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
''' A parse tree for an expression initializer.
''' </summary>
Public NotInheritable Class ExpressionInitializer
    Inherits Initializer

    Private ReadOnly _Expression As Expression

    ''' <summary>
    ''' The expression.
    ''' </summary>
    Public ReadOnly Property Expression() As Expression
        Get
            Return _Expression
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new expression initializer parse tree.
    ''' </summary>
    ''' <param name="expression">The expression.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal expression As Expression, ByVal span As Span)
        MyBase.New(TreeType.ExpressionInitializer, span)

        If expression Is Nothing Then
            Throw New ArgumentNullException("expression")
        End If

        SetParent(expression)

        _Expression = expression
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Expression)
    End Sub
End Class