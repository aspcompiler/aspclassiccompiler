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
''' A parse tree for a Resume statement.
''' </summary>
Public NotInheritable Class ResumeStatement
    Inherits LabelReferenceStatement

    Private ReadOnly _ResumeType As ResumeType
    Private ReadOnly _NextLocation As Location

    ''' <summary>
    ''' The type of the Resume statement.
    ''' </summary>
    Public ReadOnly Property ResumeType() As ResumeType
        Get
            Return _ResumeType
        End Get
    End Property

    ''' <summary>
    ''' The location of the 'Next', if any.
    ''' </summary>
    Public ReadOnly Property NextLocation() As Location
        Get
            Return _NextLocation
        End Get
    End Property

    ''' <summary>
    ''' Constructs a parse tree for a Resume statement.
    ''' </summary>
    ''' <param name="resumeType">The type of the Resume statement.</param>
    ''' <param name="nextLocation">The location of the 'Next', if any.</param>
    ''' <param name="name">The label name, if any.</param>
    ''' <param name="isLineNumber">Whether the label is a line number.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments of the parse tree.</param>
    Public Sub New(ByVal resumeType As ResumeType, ByVal nextLocation As Location, ByVal name As SimpleName, ByVal isLineNumber As Boolean, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.ResumeStatement, name, isLineNumber, span, comments)

        If resumeType < resumeType.None OrElse resumeType > resumeType.Label Then
            Throw New ArgumentOutOfRangeException("resumeType")
        End If

        _ResumeType = resumeType
        _NextLocation = nextLocation
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        If ResumeType = ResumeType.Label Then
            MyBase.GetChildTrees(childList)
        End If
    End Sub
End Class