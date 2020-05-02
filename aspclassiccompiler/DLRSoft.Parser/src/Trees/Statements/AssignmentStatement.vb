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
''' A parse tree for an assignment statement.
''' </summary>
Public NotInheritable Class AssignmentStatement
    Inherits Statement

    Private ReadOnly _TargetExpression As Expression
    Private ReadOnly _OperatorLocation As Location
    Private ReadOnly _SourceExpression As Expression
    Private ReadOnly _IsSetStatement As Boolean

    ''' <summary>
    ''' The target of the assignment.
    ''' </summary>
    Public ReadOnly Property TargetExpression() As Expression
        Get
            Return _TargetExpression
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

    Public ReadOnly Property IsSetStatement() As Boolean
        Get
            Return _IsSetStatement
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for an assignment statement.
    ''' </summary>
    ''' <param name="targetExpression">The target of the assignment.</param>
    ''' <param name="operatorLocation">The location of the operator.</param>
    ''' <param name="sourceExpression">The source of the assignment.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    ''' <param name="isSetStatement">Whether is is a set statement</param>
    Public Sub New(ByVal targetExpression As Expression, ByVal operatorLocation As Location, ByVal sourceExpression As Expression, ByVal span As Span, ByVal comments As IList(Of Comment), ByVal isSetStatement As Boolean)
        MyBase.New(TreeType.AssignmentStatement, span, comments)

        If targetExpression Is Nothing Then
            Throw New ArgumentNullException("targetExpression")
        End If

        If sourceExpression Is Nothing Then
            Throw New ArgumentNullException("sourceExpression")
        End If

        SetParent(targetExpression)
        SetParent(sourceExpression)

        _TargetExpression = targetExpression
        _OperatorLocation = operatorLocation
        _SourceExpression = sourceExpression
        _IsSetStatement = isSetStatement
    End Sub

    ''' <summary>
    ''' Constructs a new parse tree for an assignment statement.
    ''' </summary>
    ''' <param name="targetExpression">The target of the assignment.</param>
    ''' <param name="operatorLocation">The location of the operator.</param>
    ''' <param name="sourceExpression">The source of the assignment.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal targetExpression As Expression, ByVal operatorLocation As Location, ByVal sourceExpression As Expression, ByVal span As Span, ByVal comments As IList(Of Comment))
        Me.New(targetExpression, operatorLocation, sourceExpression, span, comments, False)
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, TargetExpression)
        AddChild(childList, SourceExpression)
    End Sub
End Class