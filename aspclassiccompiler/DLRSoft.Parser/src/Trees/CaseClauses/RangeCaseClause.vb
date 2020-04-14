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
''' A parse tree for a case clause that compares against a range of values.
''' </summary>
Public NotInheritable Class RangeCaseClause
    Inherits CaseClause

    Private ReadOnly _RangeExpression As Expression

    ''' <summary>
    ''' The range expression.
    ''' </summary>
    Public ReadOnly Property RangeExpression() As Expression
        Get
            Return _RangeExpression
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new range case clause parse tree.
    ''' </summary>
    ''' <param name="rangeExpression">The range expression.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal rangeExpression As Expression, ByVal span As Span)
        MyBase.New(TreeType.RangeCaseClause, span)

        If rangeExpression Is Nothing Then
            Throw New ArgumentNullException("rangeExpression")
        End If

        SetParent(rangeExpression)

        _RangeExpression = rangeExpression
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, RangeExpression)
    End Sub
End Class