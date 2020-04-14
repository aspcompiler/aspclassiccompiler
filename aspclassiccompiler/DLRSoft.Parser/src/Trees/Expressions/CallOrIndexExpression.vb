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
''' A parse tree for a call or index expression.
''' </summary>
Public NotInheritable Class CallOrIndexExpression
    Inherits Expression

    Private ReadOnly _TargetExpression As Expression
    Private ReadOnly _Arguments As ArgumentCollection

    ''' <summary>
    ''' The target of the call or index.
    ''' </summary>
    Public ReadOnly Property TargetExpression() As Expression
        Get
            Return _TargetExpression
        End Get
    End Property

    ''' <summary>
    ''' The arguments to the call or index.
    ''' </summary>
    Public ReadOnly Property Arguments() As ArgumentCollection
        Get
            Return _Arguments
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a call or index expression.
    ''' </summary>
    ''' <param name="targetExpression">The target of the call or index.</param>
    ''' <param name="arguments">The arguments to the call or index.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal targetExpression As Expression, ByVal arguments As ArgumentCollection, ByVal span As Span)
        MyBase.New(TreeType.CallOrIndexExpression, span)

        If targetExpression Is Nothing Then
            Throw New ArgumentNullException("targetExpression")
        End If

        SetParent(targetExpression)
        SetParent(arguments)

        _TargetExpression = targetExpression
        _Arguments = arguments
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, TargetExpression)
        AddChild(childList, Arguments)
    End Sub
End Class