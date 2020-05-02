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
''' A parse tree for the block of a Case Else statement.
''' </summary>
Public NotInheritable Class CaseElseBlockStatement
    Inherits BlockStatement

    Private ReadOnly _CaseElseStatement As CaseElseStatement

    ''' <summary>
    ''' The Case Else statement that started the block.
    ''' </summary>
    Public ReadOnly Property CaseElseStatement() As CaseElseStatement
        Get
            Return _CaseElseStatement
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for the block of a Case Else statement.
    ''' </summary>
    ''' <param name="caseElseStatement">The Case Else statement that started the block.</param>
    ''' <param name="statements">The statements in the block.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments of the tree.</param>
    Public Sub New(ByVal caseElseStatement As CaseElseStatement, ByVal statements As StatementCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.CaseElseBlockStatement, statements, span, comments)

        If caseElseStatement Is Nothing Then
            Throw New ArgumentNullException("caseElseStatement")
        End If

        SetParent(caseElseStatement)
        _CaseElseStatement = caseElseStatement
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, CaseElseStatement)
        MyBase.GetChildTrees(childList)
    End Sub
End Class