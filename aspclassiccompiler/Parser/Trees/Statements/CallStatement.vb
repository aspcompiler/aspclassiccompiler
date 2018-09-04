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
''' A parse tree for a method call statement.
''' </summary>
Public NotInheritable Class CallStatement
    Inherits Statement

    Private ReadOnly _CallLocation As Location
    Private ReadOnly _TargetExpression As Expression
    Private ReadOnly _Arguments As ArgumentCollection

    ''' <summary>
    ''' The location of the 'Call', if any.
    ''' </summary>
    Public ReadOnly Property CallLocation() As Location
        Get
            Return _CallLocation
        End Get
    End Property

    ''' <summary>
    ''' The target of the call.
    ''' </summary>
    Public ReadOnly Property TargetExpression() As Expression
        Get
            Return _TargetExpression
        End Get
    End Property

    ''' <summary>
    ''' The arguments to the call.
    ''' </summary>
    Public ReadOnly Property Arguments() As ArgumentCollection
        Get
            Return _Arguments
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a method call statement.
    ''' </summary>
    ''' <param name="callLocation">The location of the 'Call', if any.</param>
    ''' <param name="targetExpression">The target of the call.</param>
    ''' <param name="arguments">The arguments to the call.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments of the parse tree.</param>
    Public Sub New(ByVal callLocation As Location, ByVal targetExpression As Expression, ByVal arguments As ArgumentCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.CallStatement, span, comments)

        If targetExpression Is Nothing Then
            Throw New ArgumentNullException("targetExpression")
        End If

        SetParent(targetExpression)
        SetParent(arguments)

        _CallLocation = callLocation
        _TargetExpression = targetExpression
        _Arguments = arguments
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, TargetExpression)
        AddChild(childList, Arguments)
    End Sub
End Class