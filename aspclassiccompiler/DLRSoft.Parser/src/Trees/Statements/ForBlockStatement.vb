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
''' A parse tree for a For statement.
''' </summary>
Public NotInheritable Class ForBlockStatement
    Inherits BlockStatement

    Private ReadOnly _ControlExpression As Expression
    Private ReadOnly _ControlVariableDeclarator As VariableDeclarator
    Private ReadOnly _EqualsLocation As Location
    Private ReadOnly _LowerBoundExpression As Expression
    Private ReadOnly _ToLocation As Location
    Private ReadOnly _UpperBoundExpression As Expression
    Private ReadOnly _StepLocation As Location
    Private ReadOnly _StepExpression As Expression
    Private ReadOnly _NextStatement As NextStatement

    ''' <summary>
    ''' The control expression for the loop.
    ''' </summary>
    Public ReadOnly Property ControlExpression() As Expression
        Get
            Return _ControlExpression
        End Get
    End Property

    ''' <summary>
    ''' The control variable declarator, if any.
    ''' </summary>
    Public ReadOnly Property ControlVariableDeclarator() As VariableDeclarator
        Get
            Return _ControlVariableDeclarator
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
    ''' The lower bound of the loop.
    ''' </summary>
    Public ReadOnly Property LowerBoundExpression() As Expression
        Get
            Return _LowerBoundExpression
        End Get
    End Property

    ''' <summary>
    ''' The location of the 'To'.
    ''' </summary>
    Public ReadOnly Property ToLocation() As Location
        Get
            Return _ToLocation
        End Get
    End Property

    ''' <summary>
    ''' The upper bound of the loop.
    ''' </summary>
    Public ReadOnly Property UpperBoundExpression() As Expression
        Get
            Return _UpperBoundExpression
        End Get
    End Property

    ''' <summary>
    ''' The location of the 'Step', if any.
    ''' </summary>
    Public ReadOnly Property StepLocation() As Location
        Get
            Return _StepLocation
        End Get
    End Property

    ''' <summary>
    ''' The step of the loop, if any.
    ''' </summary>
    Public ReadOnly Property StepExpression() As Expression
        Get
            Return _StepExpression
        End Get
    End Property

    ''' <summary>
    ''' The Next statement, if any.
    ''' </summary>
    Public ReadOnly Property NextStatement() As NextStatement
        Get
            Return _NextStatement
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a For statement.
    ''' </summary>
    ''' <param name="controlExpression">The control expression for the loop.</param>
    ''' <param name="controlVariableDeclarator">The control variable declarator, if any.</param>
    ''' <param name="equalsLocation">The location of the '='.</param>
    ''' <param name="lowerBoundExpression">The lower bound of the loop.</param>
    ''' <param name="toLocation">The location of the 'To'.</param>
    ''' <param name="upperBoundExpression">The upper bound of the loop.</param>
    ''' <param name="stepLocation">The location of the 'Step', if any.</param>
    ''' <param name="stepExpression">The step of the loop, if any.</param>
    ''' <param name="statements">The statements in the For loop.</param>
    ''' <param name="nextStatement">The Next statement, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal controlExpression As Expression, ByVal controlVariableDeclarator As VariableDeclarator, ByVal equalsLocation As Location, ByVal lowerBoundExpression As Expression, ByVal toLocation As Location, ByVal upperBoundExpression As Expression, ByVal stepLocation As Location, ByVal stepExpression As Expression, ByVal statements As StatementCollection, ByVal nextStatement As NextStatement, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.ForBlockStatement, statements, span, comments)

        If controlExpression Is Nothing Then
            Throw New ArgumentNullException("controlExpression")
        End If

        SetParent(controlExpression)
        SetParent(controlVariableDeclarator)
        SetParent(lowerBoundExpression)
        SetParent(upperBoundExpression)
        SetParent(stepExpression)
        SetParent(nextStatement)

        _ControlExpression = controlExpression
        _ControlVariableDeclarator = controlVariableDeclarator
        _EqualsLocation = equalsLocation
        _LowerBoundExpression = lowerBoundExpression
        _ToLocation = toLocation
        _UpperBoundExpression = upperBoundExpression
        _StepLocation = stepLocation
        _StepExpression = stepExpression
        _NextStatement = nextStatement
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, ControlExpression)
        AddChild(childList, ControlVariableDeclarator)
        AddChild(childList, LowerBoundExpression)
        AddChild(childList, UpperBoundExpression)
        AddChild(childList, StepExpression)
        MyBase.GetChildTrees(childList)
        AddChild(childList, NextStatement)
    End Sub
End Class