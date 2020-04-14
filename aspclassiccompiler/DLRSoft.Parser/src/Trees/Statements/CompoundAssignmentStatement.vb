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
''' A parse tree for a compound assignment statement.
''' </summary>
Public NotInheritable Class CompoundAssignmentStatement
    Inherits Statement

    Private ReadOnly _TargetExpression As Expression
    Private ReadOnly _CompoundOperator As OperatorType
    Private ReadOnly _OperatorLocation As Location
    Private ReadOnly _SourceExpression As Expression

    ''' <summary>
    ''' The target of the assignment.
    ''' </summary>
    Public ReadOnly Property TargetExpression() As Expression
        Get
            Return _TargetExpression
        End Get
    End Property

    ''' <summary>
    ''' The compound operator.
    ''' </summary>
    Public ReadOnly Property CompoundOperator() As OperatorType
        Get
            Return _CompoundOperator
        End Get
    End Property

    ''' <summary>
    ''' The location of the operator.
    ''' </summary>
    Public ReadOnly Property OperatorLocation() As Location
        Get
            Return _OperatorLocation
        End Get
    End Property

    ''' <summary>
    ''' The source of the assignment.
    ''' </summary>
    Public ReadOnly Property SourceExpression() As Expression
        Get
            Return _SourceExpression
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a compound assignment statement.
    ''' </summary>
    ''' <param name="compoundOperator">The compound operator.</param>
    ''' <param name="targetExpression">The target of the assignment.</param>
    ''' <param name="operatorLocation">The location of the operator.</param>
    ''' <param name="sourceExpression">The source of the assignment.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal compoundOperator As OperatorType, ByVal targetExpression As Expression, ByVal operatorLocation As Location, ByVal sourceExpression As Expression, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.CompoundAssignmentStatement, span, comments)

        If compoundOperator < OperatorType.Plus OrElse compoundOperator > OperatorType.Power Then
            Throw New ArgumentOutOfRangeException("compoundOperator")
        End If

        If targetExpression Is Nothing Then
            Throw New ArgumentNullException("targetExpression")
        End If

        If sourceExpression Is Nothing Then
            Throw New ArgumentNullException("sourceExpression")
        End If

        SetParent(targetExpression)
        SetParent(sourceExpression)

        _CompoundOperator = compoundOperator
        _TargetExpression = targetExpression
        _OperatorLocation = operatorLocation
        _SourceExpression = sourceExpression
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, TargetExpression)
        AddChild(childList, SourceExpression)
    End Sub
End Class