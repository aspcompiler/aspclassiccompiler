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
''' A parse tree for a For Each statement.
''' </summary>
Public NotInheritable Class ForEachBlockStatement
    Inherits BlockStatement

    Private ReadOnly _EachLocation As Location
    Private ReadOnly _ControlExpression As Expression
    Private ReadOnly _ControlVariableDeclarator As VariableDeclarator
    Private ReadOnly _InLocation As Location
    Private ReadOnly _CollectionExpression As Expression
    Private ReadOnly _NextStatement As NextStatement

    ''' <summary>
    ''' The location of the 'Each'.
    ''' </summary>
    Public ReadOnly Property EachLocation() As Location
        Get
            Return _EachLocation
        End Get
    End Property

    ''' <summary>
    ''' The control expression.
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
    ''' The location of the 'In'.
    ''' </summary>
    Public ReadOnly Property InLocation() As Location
        Get
            Return _InLocation
        End Get
    End Property

    ''' <summary>
    ''' The collection expression.
    ''' </summary>
    Public ReadOnly Property CollectionExpression() As Expression
        Get
            Return _CollectionExpression
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
    ''' Constructs a new parse tree for a For Each statement.
    ''' </summary>
    ''' <param name="eachLocation">The location of the 'Each'.</param>
    ''' <param name="controlExpression">The control expression.</param>
    ''' <param name="controlVariableDeclarator">The control variable declarator, if any.</param>
    ''' <param name="inLocation">The location of the 'In'.</param>
    ''' <param name="collectionExpression">The collection expression.</param>
    ''' <param name="statements">The statements in the block.</param>
    ''' <param name="nextStatement">The Next statement, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal eachLocation As Location, ByVal controlExpression As Expression, ByVal controlVariableDeclarator As VariableDeclarator, ByVal inLocation As Location, ByVal collectionExpression As Expression, ByVal statements As StatementCollection, ByVal nextStatement As NextStatement, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.ForEachBlockStatement, statements, span, comments)

        If controlExpression Is Nothing Then
            Throw New ArgumentNullException("controlExpression")
        End If

        SetParent(controlExpression)
        SetParent(controlVariableDeclarator)
        SetParent(collectionExpression)
        SetParent(nextStatement)

        _EachLocation = eachLocation
        _ControlExpression = controlExpression
        _ControlVariableDeclarator = controlVariableDeclarator
        _InLocation = inLocation
        _CollectionExpression = collectionExpression
        _NextStatement = nextStatement
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, ControlExpression)
        AddChild(childList, ControlVariableDeclarator)
        AddChild(childList, CollectionExpression)
        MyBase.GetChildTrees(childList)
        AddChild(childList, NextStatement)
    End Sub
End Class