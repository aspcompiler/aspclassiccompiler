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
''' A parse tree for a Mid assignment statement.
''' </summary>
Public NotInheritable Class MidAssignmentStatement
    Inherits Statement

    Private ReadOnly _HasTypeCharacter As Boolean
    Private ReadOnly _LeftParenthesisLocation As Location
    Private ReadOnly _TargetExpression As Expression
    Private ReadOnly _StartCommaLocation As Location
    Private ReadOnly _StartExpression As Expression
    Private ReadOnly _LengthCommaLocation As Location
    Private ReadOnly _LengthExpression As Expression
    Private ReadOnly _RightParenthesisLocation As Location
    Private ReadOnly _OperatorLocation As Location
    Private ReadOnly _SourceExpression As Expression

    ''' <summary>
    ''' Whether the Mid identifier had a type character.
    ''' </summary>
    Public ReadOnly Property HasTypeCharacter() As Boolean
        Get
            Return _HasTypeCharacter
        End Get
    End Property

    ''' <summary>
    ''' The location of the left parenthesis.
    ''' </summary>
    Public ReadOnly Property LeftParenthesisLocation() As Location
        Get
            Return _LeftParenthesisLocation
        End Get
    End Property

    ''' <summary>
    ''' The target of the assignment.
    ''' </summary>
    Public ReadOnly Property TargetExpression() As Expression
        Get
            Return _TargetExpression
        End Get
    End Property

    ''' <summary>
    ''' The location of the comma before the start expression.
    ''' </summary>
    Public ReadOnly Property StartCommaLocation() As Location
        Get
            Return _StartCommaLocation
        End Get
    End Property

    ''' <summary>
    ''' The expression representing the start of the string to replace.
    ''' </summary>
    Public ReadOnly Property StartExpression() As Expression
        Get
            Return _StartExpression
        End Get
    End Property

    ''' <summary>
    ''' The location of the comma before the length expression, if any.
    ''' </summary>
    Public ReadOnly Property LengthCommaLocation() As Location
        Get
            Return _LengthCommaLocation
        End Get
    End Property

    ''' <summary>
    ''' The expression representing the length of the string to replace, if any.
    ''' </summary>
    Public ReadOnly Property LengthExpression() As Expression
        Get
            Return _LengthExpression
        End Get
    End Property

    ''' <summary>
    ''' The right parenthesis location.
    ''' </summary>
    Public ReadOnly Property RightParenthesisLocation() As Location
        Get
            Return _RightParenthesisLocation
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
    ''' Constructs a new parse tree for an assignment statement.
    ''' </summary>
    ''' <param name="hasTypeCharacter">Whether the Mid identifier has a type character.</param>
    ''' <param name="leftParenthesisLocation">The location of the left parenthesis.</param>
    ''' <param name="targetExpression">The target of the assignment.</param>
    ''' <param name="startCommaLocation">The location of the comma before the start expression.</param>
    ''' <param name="startExpression">The expression representing the start of the string to replace.</param>
    ''' <param name="lengthCommaLocation">The location of the comma before the length expression, if any.</param>
    ''' <param name="lengthExpression">The expression representing the length of the string to replace, if any.</param>
    ''' <param name="operatorLocation">The location of the operator.</param>
    ''' <param name="sourceExpression">The source of the assignment.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal hasTypeCharacter As Boolean, ByVal leftParenthesisLocation As Location, ByVal targetExpression As Expression, ByVal startCommaLocation As Location, ByVal startExpression As Expression, ByVal lengthCommaLocation As Location, ByVal lengthExpression As Expression, ByVal rightParenthesisLocation As Location, ByVal operatorLocation As Location, ByVal sourceExpression As Expression, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.MidAssignmentStatement, span, comments)

        If targetExpression Is Nothing Then
            Throw New ArgumentNullException("targetExpression")
        End If

        If startExpression Is Nothing Then
            Throw New ArgumentNullException("startExpression")
        End If

        If sourceExpression Is Nothing Then
            Throw New ArgumentNullException("sourceExpression")
        End If

        SetParent(targetExpression)
        SetParent(startExpression)
        SetParent(lengthExpression)
        SetParent(sourceExpression)

        _HasTypeCharacter = hasTypeCharacter
        _LeftParenthesisLocation = leftParenthesisLocation
        _TargetExpression = targetExpression
        _StartCommaLocation = startCommaLocation
        _StartExpression = startExpression
        _LengthCommaLocation = lengthCommaLocation
        _LengthExpression = lengthExpression
        _RightParenthesisLocation = rightParenthesisLocation
        _OperatorLocation = operatorLocation
        _SourceExpression = sourceExpression
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, TargetExpression)
        AddChild(childList, StartExpression)
        AddChild(childList, LengthExpression)
        AddChild(childList, SourceExpression)
    End Sub
End Class