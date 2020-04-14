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
''' A parse tree for a block statement.
''' </summary>
Public MustInherit Class BlockStatement
    Inherits Statement

    Private ReadOnly _Statements As StatementCollection

    ''' <summary>
    ''' The statements in the block.
    ''' </summary>
    Public ReadOnly Property Statements() As StatementCollection
        Get
            Return _Statements
        End Get
    End Property

    Protected Sub New(ByVal type As TreeType, ByVal statements As StatementCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(type, span, comments)

        _Statements = statements
        SetParent(statements)
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Statements)
    End Sub
End Class