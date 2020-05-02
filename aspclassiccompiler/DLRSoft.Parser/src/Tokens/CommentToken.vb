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
''' A comment token.
''' </summary>
Public NotInheritable Class CommentToken
    Inherits Token

    Private ReadOnly _IsREM As Boolean    ' Was the comment preceded by a quote or by REM?
    Private ReadOnly _Comment As String   ' Comment can be Nothing

    ''' <summary>
    ''' Whether the comment was preceded by REM.
    ''' </summary>
    Public ReadOnly Property IsREM() As Boolean
        Get
            Return _IsREM
        End Get
    End Property

    ''' <summary>
    ''' The text of the comment.
    ''' </summary>
    Public ReadOnly Property Comment() As String
        Get
            Return _Comment
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new comment token.
    ''' </summary>
    ''' <param name="comment">The comment value.</param>
    ''' <param name="isREM">Whether the comment was preceded by REM.</param>
    ''' <param name="span">The location of the comment.</param>
    Public Sub New(ByVal comment As String, ByVal isREM As Boolean, ByVal span As Span)
        MyBase.New(TokenType.Comment, span)

        If comment Is Nothing Then
            Throw New ArgumentNullException("comment")
        End If

        _IsREM = isREM
        _Comment = comment
    End Sub
End Class