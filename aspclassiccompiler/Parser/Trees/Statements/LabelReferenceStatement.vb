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
''' A parse tree for a statement that refers to a label.
''' </summary>
Public MustInherit Class LabelReferenceStatement
    Inherits Statement

    Private ReadOnly _Name As SimpleName
    Private ReadOnly _IsLineNumber As Boolean

    ''' <summary>
    ''' The name of the label being referred to.
    ''' </summary>
    Public ReadOnly Property Name() As SimpleName
        Get
            Return _Name
        End Get
    End Property

    ''' <summary>
    ''' Whether the label is a line number.
    ''' </summary>
    Public ReadOnly Property IsLineNumber() As Boolean
        Get
            Return _IsLineNumber
        End Get
    End Property

    Protected Sub New(ByVal type As TreeType, ByVal name As SimpleName, ByVal isLineNumber As Boolean, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(type, span, comments)

        Debug.Assert(type = TreeType.GotoStatement OrElse type = TreeType.LabelStatement OrElse _
                     type = TreeType.OnErrorStatement OrElse type = TreeType.ResumeStatement)

        SetParent(name)
        _Name = name
        _IsLineNumber = isLineNumber
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Name)
    End Sub
End Class