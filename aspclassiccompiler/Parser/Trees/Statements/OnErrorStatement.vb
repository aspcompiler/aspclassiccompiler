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
''' A parse tree for an On Error statement.
''' </summary>
Public NotInheritable Class OnErrorStatement
    Inherits LabelReferenceStatement

    Private ReadOnly _OnErrorType As OnErrorType
    Private ReadOnly _ErrorLocation As Location
    Private ReadOnly _ResumeOrGoToLocation As Location
    Private ReadOnly _NextOrZeroOrMinusLocation As Location
    Private ReadOnly _OneLocation As Location

    ''' <summary>
    ''' The type of On Error statement.
    ''' </summary>
    Public ReadOnly Property OnErrorType() As OnErrorType
        Get
            Return _OnErrorType
        End Get
    End Property

    ''' <summary>
    ''' The location of the 'Error'.
    ''' </summary>
    Public ReadOnly Property ErrorLocation() As Location
        Get
            Return _ErrorLocation
        End Get
    End Property

    ''' <summary>
    ''' The location of the 'Resume' or 'GoTo'.
    ''' </summary>
    Public ReadOnly Property ResumeOrGoToLocation() As Location
        Get
            Return _ResumeOrGoToLocation
        End Get
    End Property

    ''' <summary>
    ''' The location of the 'Next', '0' or '-', if any.
    ''' </summary>
    Public ReadOnly Property NextOrZeroOrMinusLocation() As Location
        Get
            Return _NextOrZeroOrMinusLocation
        End Get
    End Property

    ''' <summary>
    ''' The location of the '1', if any.
    ''' </summary>
    Public ReadOnly Property OneLocation() As Location
        Get
            Return _OneLocation
        End Get
    End Property

    ''' <summary>
    ''' Constructs a parse tree for an On Error statement.
    ''' </summary>
    ''' <param name="onErrorType">The type of the On Error statement.</param>
    ''' <param name="errorLocation">The location of the 'Error'.</param>
    ''' <param name="resumeOrGoToLocation">The location of the 'Resume' or 'GoTo'.</param>
    ''' <param name="nextOrZeroOrMinusLocation">The location of the 'Next', '0' or '-', if any.</param>
    ''' <param name="oneLocation">The location of the '1', if any.</param>
    ''' <param name="name">The label to branch to, if any.</param>
    ''' <param name="isLineNumber">Whether the label is a line number.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal onErrorType As OnErrorType, ByVal errorLocation As Location, ByVal resumeOrGoToLocation As Location, ByVal nextOrZeroOrMinusLocation As Location, ByVal oneLocation As Location, ByVal name As SimpleName, ByVal isLineNumber As Boolean, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.OnErrorStatement, name, isLineNumber, span, comments)

        If onErrorType < onErrorType.Bad OrElse onErrorType > onErrorType.Label Then
            Throw New ArgumentOutOfRangeException("onErrorType")
        End If

        _OnErrorType = onErrorType
        _ErrorLocation = errorLocation
        _ResumeOrGoToLocation = resumeOrGoToLocation
        _NextOrZeroOrMinusLocation = nextOrZeroOrMinusLocation
        _OneLocation = oneLocation
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        If OnErrorType = OnErrorType.Label Then
            MyBase.GetChildTrees(childList)
        End If
    End Sub
End Class