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
''' A parse tree for a Case statement.
''' </summary>
Public NotInheritable Class CaseStatement
    Inherits Statement

    Private ReadOnly _CaseClauses As CaseClauseCollection

    ''' <summary>
    ''' The clauses in the Case statement.
    ''' </summary>
    Public ReadOnly Property CaseClauses() As CaseClauseCollection
        Get
            Return _CaseClauses
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a Case statement.
    ''' </summary>
    ''' <param name="caseClauses">The clauses in the Case statement.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments on the parse tree.</param>
    Public Sub New(ByVal caseClauses As CaseClauseCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.CaseStatement, span, comments)

        If caseClauses Is Nothing Then
            Throw New ArgumentNullException("caseClauses")
        End If

        SetParent(caseClauses)
        _CaseClauses = caseClauses
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, CaseClauses)
    End Sub
End Class