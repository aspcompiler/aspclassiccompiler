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
''' A parse tree for the block of a Case statement.
''' </summary>
Public NotInheritable Class CaseBlockStatement
    Inherits BlockStatement

    Private ReadOnly _CaseStatement As CaseStatement

    ''' <summary>
    ''' The Case statement that started the block.
    ''' </summary>
    Public ReadOnly Property CaseStatement() As CaseStatement
        Get
            Return _CaseStatement
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for the block of a Case statement.
    ''' </summary>
    ''' <param name="caseStatement">The Case statement that started the block.</param>
    ''' <param name="statements">The statements in the block.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments of the tree.</param>
    Public Sub New(ByVal caseStatement As CaseStatement, ByVal statements As StatementCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.CaseBlockStatement, statements, span, comments)

        If caseStatement Is Nothing Then
            Throw New ArgumentNullException("caseStatement")
        End If

        SetParent(caseStatement)
        _CaseStatement = caseStatement
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, CaseStatement)
        MyBase.GetChildTrees(childList)
    End Sub
End Class