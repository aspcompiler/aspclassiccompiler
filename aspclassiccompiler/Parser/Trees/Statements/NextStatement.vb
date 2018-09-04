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
''' A parse tree for a Next statement.
''' </summary>
Public NotInheritable Class NextStatement
    Inherits Statement

    Private ReadOnly _Variables As ExpressionCollection

    ''' <summary>
    ''' The loop control variables.
    ''' </summary>
    Public ReadOnly Property Variables() As ExpressionCollection
        Get
            Return _Variables
        End Get
    End Property

    ''' <summary>
    ''' Constructs a parse tree for a Next statement.
    ''' </summary>
    ''' <param name="variables">The loop control variables.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal variables As ExpressionCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.NextStatement, span, comments)

        SetParent(variables)
        _Variables = variables
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Variables)
    End Sub
End Class