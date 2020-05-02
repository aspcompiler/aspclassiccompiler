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
''' A parse tree for an Erase statement.
''' </summary>
Public NotInheritable Class EraseStatement
    Inherits Statement

    Private ReadOnly _Variables As ExpressionCollection

    ''' <summary>
    ''' The variables to erase.
    ''' </summary>
    Public ReadOnly Property Variables() As ExpressionCollection
        Get
            Return _Variables
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for an Erase statement.
    ''' </summary>
    ''' <param name="variables">The variables to erase.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal variables As ExpressionCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.EraseStatement, span, comments)

        If variables Is Nothing Then
            Throw New ArgumentNullException("variables")
        End If

        SetParent(variables)
        _Variables = variables
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Variables)
    End Sub
End Class