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
''' A parse tree for a comment.
''' </summary>
Public NotInheritable Class Comment
    Inherits Tree

    Private ReadOnly _Comment As String
    Private ReadOnly _IsREM As Boolean

    ''' <summary>
    ''' The text of the comment.
    ''' </summary>
    Public ReadOnly Property Comment() As String
        Get
            Return _Comment
        End Get
    End Property

    ''' <summary>
    ''' Whether the comment is a REM comment.
    ''' </summary>
    Public ReadOnly Property IsREM() As Boolean
        Get
            Return _IsREM
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new comment parse tree.
    ''' </summary>
    ''' <param name="comment">The text of the comment.</param>
    ''' <param name="isREM">Whether the comment is a REM comment.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal comment As String, ByVal isREM As Boolean, ByVal span As Span)
        MyBase.New(TreeType.Comment, span)

        If comment Is Nothing Then
            Throw New ArgumentNullException("comment")
        End If

        _Comment = comment
        _IsREM = isREM
    End Sub
End Class