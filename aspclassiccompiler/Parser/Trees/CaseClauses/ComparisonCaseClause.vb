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
''' A parse tree for a case clause that compares values.
''' </summary>
Public NotInheritable Class ComparisonCaseClause
    Inherits CaseClause

    Private ReadOnly _IsLocation As Location
    Private ReadOnly _ComparisonOperator As OperatorType
    Private ReadOnly _OperatorLocation As Location
    Private ReadOnly _Operand As Expression

    ''' <summary>
    ''' The location of the 'Is', if any.
    ''' </summary>
    Public ReadOnly Property IsLocation() As Location
        Get
            Return _IsLocation
        End Get
    End Property

    ''' <summary>
    ''' The comparison operator used in the case clause.
    ''' </summary>
    Public ReadOnly Property ComparisonOperator() As OperatorType
        Get
            Return _ComparisonOperator
        End Get
    End Property

    ''' <summary>
    ''' The location of the comparison operator.
    ''' </summary>
    Public ReadOnly Property OperatorLocation() As Location
        Get
            Return _OperatorLocation
        End Get
    End Property

    ''' <summary>
    ''' The operand of the case clause.
    ''' </summary>
    Public ReadOnly Property Operand() As Expression
        Get
            Return _Operand
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a comparison case clause.
    ''' </summary>
    ''' <param name="isLocation">The location of the 'Is', if any.</param>
    ''' <param name="comparisonOperator">The comparison operator used.</param>
    ''' <param name="operatorLocation">The location of the comparison operator.</param>
    ''' <param name="operand">The operand of the comparison.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal isLocation As Location, ByVal comparisonOperator As OperatorType, ByVal operatorLocation As Location, ByVal operand As Expression, ByVal span As Span)
        MyBase.New(TreeType.ComparisonCaseClause, span)

        If operand Is Nothing Then
            Throw New ArgumentNullException("operand")
        End If

        If comparisonOperator < OperatorType.Equals OrElse comparisonOperator > OperatorType.GreaterThanEquals Then
            Throw New ArgumentOutOfRangeException("comparisonOperator")
        End If

        SetParent(operand)

        _IsLocation = isLocation
        _ComparisonOperator = comparisonOperator
        _OperatorLocation = operatorLocation
        _Operand = operand
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Operand)
    End Sub
End Class